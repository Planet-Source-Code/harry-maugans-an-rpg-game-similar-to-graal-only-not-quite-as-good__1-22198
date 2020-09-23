VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMainOnline 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game (Online Version 1.0)"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9615
   Icon            =   "frmMainOnline.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   448
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Sock 
      Left            =   0
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrSword 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1920
      Top             =   0
   End
   Begin VB.Timer tmrHit 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer tmrIT 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer tmrOT 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame frmMenu 
      BackColor       =   &H00000000&
      Caption         =   "Menu (TAB)"
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Left            =   0
      TabIndex        =   4
      Top             =   5160
      Visible         =   0   'False
      Width           =   9615
      Begin VB.TextBox txtBombs 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtArrows 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Unnamed"
         Top             =   360
         Width           =   855
      End
      Begin VB.Frame frmActiveWeapons 
         BackColor       =   &H00404040&
         Caption         =   "Active Weapons"
         ForeColor       =   &H0080FFFF&
         Height          =   1095
         Left            =   6000
         TabIndex        =   16
         Top             =   360
         Width           =   3375
         Begin VB.Image imgIW 
            Height          =   495
            Left            =   2160
            Stretch         =   -1  'True
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Inventory Weapon:"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1680
            TabIndex        =   19
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   """S"": Weapon"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   """A"": Sword"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox txtIT 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtOT 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtGold 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtTotalHealth 
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Text            =   "10"
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtHealth 
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Text            =   "10"
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtHealthDSP 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "10/10"
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Bombs:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   25
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Arrows:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   23
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Idle Time:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Online Time:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Gold:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Health:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame frmWeapons 
      BackColor       =   &H00000000&
      Caption         =   "Weapons (~)"
      ForeColor       =   &H0000FF00&
      Height          =   6720
      Left            =   7920
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Image imgWeapon 
         Height          =   495
         Index           =   14
         Left            =   240
         Stretch         =   -1  'True
         Top             =   4560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgWeapon 
         Height          =   495
         Index           =   13
         Left            =   960
         Stretch         =   -1  'True
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgWeapon 
         Height          =   495
         Index           =   12
         Left            =   240
         Stretch         =   -1  'True
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgWeapon 
         Height          =   495
         Index           =   11
         Left            =   960
         Stretch         =   -1  'True
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgWeapon 
         Height          =   495
         Index           =   10
         Left            =   240
         Stretch         =   -1  'True
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgWeapon 
         Height          =   495
         Index           =   9
         Left            =   960
         Stretch         =   -1  'True
         Top             =   2760
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgWeapon 
         Height          =   495
         Index           =   8
         Left            =   240
         Stretch         =   -1  'True
         Top             =   2760
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgWeapon 
         Height          =   495
         Index           =   7
         Left            =   960
         Stretch         =   -1  'True
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgWeapon 
         Height          =   495
         Index           =   6
         Left            =   240
         Stretch         =   -1  'True
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgWeapon 
         Height          =   495
         Index           =   5
         Left            =   960
         Stretch         =   -1  'True
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgWeapon 
         Height          =   495
         Index           =   4
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgWeapon 
         Height          =   495
         Index           =   3
         Left            =   960
         Stretch         =   -1  'True
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgWeapon 
         Height          =   495
         Index           =   2
         Left            =   240
         Stretch         =   -1  'True
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgWeapon 
         Height          =   495
         Index           =   1
         Left            =   960
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgWeapon 
         Height          =   495
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Timer tmrName 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1440
      Top             =   0
   End
   Begin VB.ListBox lstNoGo 
      Height          =   1425
      ItemData        =   "frmMainOnline.frx":000C
      Left            =   720
      List            =   "frmMainOnline.frx":000E
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtCommands 
      Height          =   375
      Left            =   -2.45760e5
      TabIndex        =   0
      Top             =   7875
      Width           =   495
   End
   Begin VB.Image imgBadGuy 
      Height          =   495
      Index           =   0
      Left            =   5280
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgBush 
      Height          =   465
      Index           =   0
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image imgBomb 
      Height          =   465
      Index           =   0
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image imgArrow 
      Height          =   465
      Index           =   0
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image img3 
      Height          =   855
      Left            =   2520
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgGold 
      Height          =   465
      Index           =   0
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image imgSword 
      Height          =   480
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMiniMap 
      Height          =   1575
      Left            =   6120
      Picture         =   "frmMainOnline.frx":0010
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin ComctlLib.ImageList imlFloorTiles 
      Left            =   5760
      Top             =   120
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
            Picture         =   "frmMainOnline.frx":1685
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":22D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":2F29
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":3B7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":47CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":541F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":6071
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":6CC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":7915
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":8567
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":91B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":9E0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":AA5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":B6AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":C301
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":CF53
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":DBA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":E7F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":F449
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":1009B
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":10CED
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":1193F
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":12591
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":131E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":13E35
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":14A87
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":156D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":1632B
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":16B7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":177CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":18421
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":19073
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":19CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":1A917
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":1B569
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":1C1BB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgNPC 
      Height          =   495
      Index           =   0
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.Image img2 
      Height          =   615
      Left            =   3600
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image img 
      Height          =   375
      Left            =   3840
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Pic 
      Height          =   375
      Left            =   2400
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image PlayerImage 
      Height          =   690
      Index           =   0
      Left            =   600
      ToolTipText     =   "You"
      Top             =   1800
      Width           =   705
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   480
      Index           =   0
      Left            =   2280
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin ComctlLib.ImageList imlGuy 
      Left            =   4320
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   20
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":1CE0D
            Key             =   "Up1"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":1E95F
            Key             =   "Up2"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":204B1
            Key             =   "Up3"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":22003
            Key             =   "Up4"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":23B55
            Key             =   "Down1"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":256A7
            Key             =   "Down2"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":271F9
            Key             =   "Down3"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":28D4B
            Key             =   "Down4"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":2A89D
            Key             =   "Left1"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":2C3EF
            Key             =   "Left2"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":2DF41
            Key             =   "Left3"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":2FA93
            Key             =   "Left4"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":315E5
            Key             =   "Right1"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":33137
            Key             =   "Right2"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":34C89
            Key             =   "Right3"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":367DB
            Key             =   "Right4"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":3832D
            Key             =   "UpSword"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":39E7F
            Key             =   "DownSword"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":3B9D1
            Key             =   "LeftSword"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":3D523
            Key             =   "RightSword"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlPeople 
      Left            =   5040
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":3F075
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":420C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":45119
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":4816B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":4B1BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":4E20F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":51261
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":542B3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlWeapons 
      Left            =   6480
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":57305
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":58E57
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":5A9A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":5C4FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":5E04D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":5FB9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":616F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":63243
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":64D95
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":668E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":68439
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":69F8B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlStatus 
      Left            =   7200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":6BADD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":6C72F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":6D381
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":6DFD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":6EC25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":6F877
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":704C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":7111B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":71F3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":72B8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":737E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMainOnline.frx":74433
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblName2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   180
      TabIndex        =   3
      Top             =   180
      Visible         =   0   'False
      Width           =   6000
   End
End
Attribute VB_Name = "frmMainOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Const TileWidth = 32
Const TileHeight = 32
Public Scrolling As Boolean
Public SignData As String
Dim OnlineHours As String
Dim OnlineMinutes As String
Dim OnlineSeconds As String
Dim IdleHours As String
Dim IdleMinutes As String
Dim IdleSeconds As String
Dim HitTmr As Integer
Public Dir As Integer
Public Dir2 As Integer
Public PX3 As Integer
Public PY3 As Integer

Private Sub Form_GotFocus()

    frmMainOnline.Height = Player.Map.MapHeight * 32
    frmMainOnline.Width = Player.Map.MapWidth * 32
    
End Sub

Public Function SoundCardInstalled() As Boolean
     SoundCardInstalled = waveOutGetNumDevs > 0
End Function

Private Sub Form_Load()
    Dim Result&
    Dim i As Integer
    
    frmMainOnline.BackColor = vbBlack
    
    Load frmDummy2
       
    Load_Player
    Warp frmDB.txtPlayerX.Text, frmDB.txtPlayerY.Text, frmDB.txtLevel.Text
    Load_Weapons
    Load_Resistances
    Load_Sword
    
    img.Width = Me.Width
    img.Height = Me.Height
    img2.Width = Me.Width
    img2.Height = Me.Height
    img3.Width = Me.Width
    img3.Height = Me.Height
    
    Player.PlayerX = PlayerImage(0).Left / 32
    Player.PlayerY = PlayerImage(0).Top / 32
    Player.MapPos = Player.PlayerX & "," & Player.PlayerY
    
    xSprite = 0

    frmMusic.Show
    frmMusic.Visible = False
    ModFormMapPos.Init
    
    frmMenu.Top = frmMenu.Top + frmMenu.Height
    frmWeapons.Left = frmWeapons.Left + frmWeapons.Width
    txtHealth.Text = frmDB.txtHealth.Text
    txtTotalHealth.Text = frmDB.txtTotalHealth.Text
    HitTmr = "500"
    Dir = 2
End Sub

Public Function Load_Sword()
    Dim FoundIt As Boolean
    Dim Temp, Temp2, Temp3 As String

    FoundIt = False
        
    If frmDB.txtSword.Text = "" Or frmDB.txtSword.Text = Chr(34) & " " & Chr(34) Then frmDB.txtSword = "Newbie-Sword"
        Open App.Path & "\Inventory\Swords.txt" For Input As #1
        Do Until EOF(1)
            Line Input #1, Temp
            Open App.Path & "\Temp4.txt" For Output As #2
                Print #2, Temp
            Close #2
            Open App.Path & "\Temp4.txt" For Input As #2
                Input #2, Temp
                Input #2, Temp2
                Input #2, Temp3
                If Temp & "," & Temp3 = frmDB.txtSword.Text Then
                    Player.AttackPower = Temp3
                    Input #2, Temp
                    Input #2, Temp
                    frmGraphics.imgSwordUp.Picture = LoadPicture(App.Path & "\Images\Swords\" & Temp)
                    Input #2, Temp
                    frmGraphics.imgSwordDown.Picture = LoadPicture(App.Path & "\Images\Swords\" & Temp)
                    Input #2, Temp
                    frmGraphics.imgSwordRight.Picture = LoadPicture(App.Path & "\Images\Swords\" & Temp)
                    Input #2, Temp
                    frmGraphics.imgSwordLeft.Picture = LoadPicture(App.Path & "\Images\Swords\" & Temp)
                    FoundIt = True
                End If
            Close #2
            Kill App.Path & "\Temp4.txt"
            If FoundIt = True Then GoTo mkay
        Loop
mkay:
        Close #1
End Function

Public Function Load_Resistances()
    If frmIS.imgShield.Tag = "" Or frmIS.imgShield.Tag = "0" Then GoTo NoShield
        Player.Evasion = frmIS.imgShield.Tag
NoShield:
    If frmIS.imgArmor.Tag = "" Or frmIS.imgArmor.Tag = "0" Then GoTo NoArmor
        Player.Defense = frmIS.imgArmor.Tag
NoArmor:
    If frmIS.imgHelmet.Tag = "" Or frmIS.imgHelmet.Tag = "0" Then GoTo NoHelmet
        Player.Defense = Player.Defense + frmIS.imgHelmet.Tag
NoHelmet:
    If frmIS.imgSword.Tag = "" Or frmIS.imgSword.Tag = "0" Then GoTo NoSword
        Player.AttackPower = frmIS.imgSword.Tag
NoSword:
End Function

Public Function Load_Weapons()
    Dim i As Integer
    Dim Data As String
    
    i = 0
    Do Until i = 15
        If frmDB.lstWeapons.List(i) <> "" And frmDB.lstWeapons.List(i) <> "none" Then
            Open App.Path & "\Temp2.txt" For Output As #7
                Print #7, frmDB.lstWeapons.List(i)
            Close #7
            Open App.Path & "\Temp2.txt" For Input As #7
                Input #7, Data
                imgWeapon(i).ToolTipText = Data
                Input #7, Data
                imgWeapon(i).Picture = LoadPicture(App.Path & "\Images\" & Data)
                Input #7, Data
                imgWeapon(i).Tag = Data
            Close #7
                imgWeapon(i).Visible = True
                Kill App.Path & "\Temp2.txt"
        End If
        i = i + 1
    Loop
        If frmDB.txtActiveWeapon.Text >= 1 Then
            imgWeapon_Click (frmDB.txtActiveWeapon.Text - 1)
        End If
End Function

Public Function DisplayMapLine(TLine As Integer)
    Dim i As Integer
    Dim p As Integer
    Dim Num As String
    Dim x As Integer

    i = 1
    x = 0
    Do Until i >= Len(Player.Map.MapData) + 1
        Num = ""
    Do
        If Mid(Player.Map.MapData, i, 1) = "," Or i >= Len(Player.Map.MapData) + 1 Then GoTo Next1
        Num = Num & Mid(Player.Map.MapData, i, 1)
        i = i + 1
    Loop
Next1:
    p = Num
    Set Pic.Picture = imlFloorTiles.ListImages(p).Picture
    frmDummy2.PaintPicture Pic.Picture, x * 32, TLine * 32
    
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
    Do Until i >= Len(Player.Map.MapData) + 1
        Num = ""
    Do
        If Mid(Player.Map.MapData, i, 1) = "," Or i >= Len(Player.Map.MapData) + 1 Then GoTo Next1
        Num = Num & Mid(Player.Map.MapData, i, 1)
        i = i + 1
    Loop
Next1:
    p = Num
    Set Pic.Picture = imlFloorTiles.ListImages(p).Picture
    frmMainOnline.PaintPicture Pic.Picture, x * 32, TLine * 32
    
    i = i + 1
    x = x + 1
    
    Loop
    
    
End Function
Private Sub Form_Paint()
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDB.SaveData3
    If Player.Direction.Dead = False Then End
End Sub

Public Function Load_Player()
    Dim i As Integer

    'set the player type
    GoingDownOnline
    Player.Speed = "32"
    Player.Health = frmDB.txtHealth.Text
    Player.Arrows = frmDB.txtArrows.Text
    txtArrows.Text = frmDB.txtArrows.Text
    Player.Bombs = frmDB.txtBombs.Text
    txtBombs.Text = frmDB.txtBombs.Text
    Player.Gold = frmDB.txtGold.Text
    Player.PlayerX = PlayerImage(0).Left
    Player.PlayerY = PlayerImage(0).Top
    txtHealth.Text = frmDB.txtHealth.Text
    txtTotalHealth.Text = frmDB.txtTotalHealth.Text
    txtGold.Text = frmDB.txtGold.Text
    txtName.Text = Player.Name
    PX3 = frmDB.txtPlayerX.Text
    PY3 = frmDB.txtPlayerY.Text
    
    
    i = 1
    OnlineHours = ""
    OnlineMinutes = ""
    OnlineSeconds = ""
    Do
        If Mid(frmDB.txtOnlineTime.Text, i, 1) = ":" Then GoTo DoneOne
        OnlineHours = OnlineHours & Mid(frmDB.txtOnlineTime.Text, i, 1)
        i = i + 1
    Loop
DoneOne:
    i = i + 1
    Do
        If Mid(frmDB.txtOnlineTime.Text, i, 1) = ":" Then GoTo DoneTwo
        OnlineMinutes = OnlineMinutes & Mid(frmDB.txtOnlineTime.Text, i, 1)
        i = i + 1
    Loop
DoneTwo:
    i = i + 1
    Do
        If i > Len(frmDB.txtOnlineTime.Text) Then GoTo DoneThree
        OnlineSeconds = OnlineSeconds & Mid(frmDB.txtOnlineTime.Text, i, 1)
        i = i + 1
    Loop
DoneThree:
    tmrOT.Enabled = True
    
    
    
    
        i = 1
    IdleHours = ""
    IdleMinutes = ""
    IdleSeconds = ""
    Do
        If Mid(frmDB.txtIdleTime.Text, i, 1) = ":" Then GoTo DoneOne2
        IdleHours = IdleHours & Mid(frmDB.txtIdleTime.Text, i, 1)
        i = i + 1
    Loop
DoneOne2:
    i = i + 1
    Do
        If Mid(frmDB.txtIdleTime.Text, i, 1) = ":" Then GoTo DoneTwo2
        IdleMinutes = IdleMinutes & Mid(frmDB.txtIdleTime.Text, i, 1)
        i = i + 1
    Loop
DoneTwo2:
    i = i + 1
    Do
        If i > Len(frmDB.txtIdleTime.Text) Then GoTo DoneThree2
        IdleSeconds = IdleSeconds & Mid(frmDB.txtIdleTime.Text, i, 1)
        i = i + 1
    Loop
DoneThree2:
    tmrIT.Enabled = True

End Function


Private Sub frmMainOnline_GotFocus()
    frmMainOnline.Refresh
    txtCommands.SetFocus
End Sub

Private Sub imgWeapon_Click(Index As Integer)
    imgIW.Picture = imgWeapon(Index).Picture
    imgIW.ToolTipText = imgWeapon(Index).Picture
    imgIW.Tag = imgWeapon(Index).Tag
    frmDB.txtActiveWeapon.Text = Index + 1
End Sub

Private Sub PlayerImage_Click(Index As Integer)
    txtCommands.SetFocus
End Sub


Private Sub tmrHit_Timer()
    If tmrHit.Enabled = True Then
        If tmrHit.Interval >= HitTmr Then
            If PlayerImage(0).Visible = True Then
                PlayerImage(0).Visible = False
            Else
                PlayerImage(0).Visible = True
            End If
            HitTmr = HitTmr - 50
            tmrHit.Interval = HitTmr
            If HitTmr = 0 Then
                HitTmr = 500
                tmrHit.Interval = HitTmr
                PlayerImage(0).Visible = True
                tmrHit.Enabled = False
            End If
            DoEvents
        End If
    End If
End Sub

Private Sub tmrIT_Timer()
    If tmrIT.Interval = "1000" Then
        If Player.Direction.Stopped = True Then
            If IdleSeconds = "59" Then GoTo NextMinute2
                IdleSeconds = IdleSeconds + 1
                If Len(IdleSeconds) = 1 Then IdleSeconds = "0" & IdleSeconds
                If Len(IdleMinutes) = 1 Then IdleMinutes = "0" & IdleMinutes
            GoTo It2
NextMinute2:
            If IdleMinutes = "59" Then GoTo NextHour2
                IdleMinutes = IdleMinutes + 1
                If Len(IdleMinutes) = 1 Then IdleMinutes = "0" & IdleMinutes
                IdleSeconds = "00"
            GoTo It2
NextHour2:
            IdleHours = IdleHours + 1
            IdleMinutes = "00"
            IdleSeconds = "00"
            GoTo It2
It2:
            txtIT.Text = IdleHours & ":" & IdleMinutes & ":" & IdleSeconds
        Else
            IdleSeconds = "0": IdleMinutes = "0": IdleHours = "0"
            txtIT.Text = "0:00:00"
        End If
    End If

End Sub

Private Sub tmrName_Timer()
    If tmrName.Interval = "2000" Then
        lblName.Visible = False
        lblName2.Visible = False
        lblName.Caption = ""
        lblName2.Caption = ""
        tmrName.Enabled = False
    End If
End Sub

Private Sub tmrOT_Timer()
    If tmrOT.Interval = "1000" Then
        If OnlineSeconds = "59" Then GoTo NextMinute
            OnlineSeconds = OnlineSeconds + 1
            If Len(OnlineSeconds) = 1 Then OnlineSeconds = "0" & OnlineSeconds
            If Len(OnlineMinutes) = 1 Then OnlineMinutes = "0" & OnlineMinutes
        GoTo it
NextMinute:
        If OnlineMinutes = "59" Then GoTo NextHour
            OnlineMinutes = OnlineMinutes + 1
            If Len(OnlineMinutes) = 1 Then OnlineMinutes = "0" & OnlineMinutes
            OnlineSeconds = "00"
        GoTo it
NextHour:
        OnlineHours = OnlineHours + 1
        OnlineMinutes = "00"
        OnlineSeconds = "00"
        GoTo it
it:
    txtOT.Text = OnlineHours & ":" & OnlineMinutes & ":" & OnlineSeconds
    frmDB.txtOnlineTime.Text = txtOT.Text
    End If
End Sub

Private Sub tmrSword_Timer()
    If tmrSword.Interval = 500 Then
        imgSword.Visible = False
        Scrolling = False
        tmrSword.Enabled = False
    End If
End Sub

Private Sub txtCommands_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    If Scrolling = True Then Exit Sub
    Dir2 = Dir
    If KeyCode = vbKeyUp Then GoingUpOnline
    If KeyCode = vbKeyDown Then GoingDownOnline
    If KeyCode = vbKeyLeft Then GoingLeftOnline
    If KeyCode = vbKeyRight Then GoingRightOnline

    If Player.Direction.Stopped = False Then IdleSeconds = "0": IdleMinutes = "0": IdleHours = "0"
    
    If Player.PlayerX = 19 And KeyCode = vbKeyRight Then
        CheckWallPos
        Exit Sub
    End If
    If Player.PlayerX = 0 And KeyCode = vbKeyLeft Then
        CheckWallPos
        Exit Sub
    End If
    If Player.PlayerY = 0 And KeyCode = vbKeyUp Then
        CheckWallPos
        Exit Sub
    End If
    If Player.PlayerY = 13 And KeyCode = vbKeyDown Then
        CheckWallPos
        Exit Sub
    End If
    
    If Player.Direction.Up = True Then
        If CheckForSigns = True Then
            frmSign.Show vbModal, Me
            txtCommands.SetFocus
        End If
    End If
    If Player.Direction.Stopped = False Then IdleSeconds = "0": IdleMinutes = "0": IdleHours = "0"
    
    If CheckForWalls(KeyCode) = False Then GoTo IsAWall
        
        If Player.Direction.Stopped = False Then
            If KeyCode = vbKeyLeft Then
                If Right((PlayerImage(0).Left / 32), 2) <> ".5" And Player.PlayerX >= 0 And Player.PlayerX <= 20 And Player.PlayerY >= 0 And Player.PlayerY <= 14 Then
                    Player.PlayerX = Player.PlayerX - 1
                    frmDB.txtPlayerX.Text = Player.PlayerX
                End If
                PlayerImage(0).Left = PlayerImage(0).Left - Player.Speed
            End If
            If KeyCode = vbKeyRight Then
                If Right((PlayerImage(0).Left / 32), 2) <> ".5" And Player.PlayerX >= 0 And Player.PlayerX <= 20 And Player.PlayerY >= 0 And Player.PlayerY <= 14 Then
                    Player.PlayerX = Player.PlayerX + 1
                    frmDB.txtPlayerX.Text = Player.PlayerX
                End If
                PlayerImage(0).Left = PlayerImage(0).Left + Player.Speed
            End If
            If KeyCode = vbKeyUp Then
                If Right((PlayerImage(0).Top / 32), 2) <> ".5" And Player.PlayerX >= 0 And Player.PlayerX <= 20 And Player.PlayerY >= 0 And Player.PlayerY <= 14 Then
                    Player.PlayerY = Player.PlayerY - 1
                    frmDB.txtPlayerY.Text = Player.PlayerY
                End If
                PlayerImage(0).Top = PlayerImage(0).Top - Player.Speed
            End If
            If KeyCode = vbKeyDown Then
                If Right((PlayerImage(0).Top / 32), 2) <> ".5" And Player.PlayerX >= 0 And Player.PlayerX <= 20 And Player.PlayerY >= 0 And Player.PlayerY <= 14 Then
                    Player.PlayerY = Player.PlayerY + 1
                    frmDB.txtPlayerY.Text = Player.PlayerY
                End If
                PlayerImage(0).Top = PlayerImage(0).Top + Player.Speed
            End If
IsAWall:
            
            If KeyCode = vbKeyA Then
                Scrolling = True
                PlayWav "Sword.wav"
                If Dir = 1 Then
                    imgSword.Top = PlayerImage(0).Top - PlayerImage(0).Height
                    imgSword.Left = PlayerImage(0).Left
                    imgSword.Picture = frmGraphics.imgSwordUp.Picture
                End If
                If Dir = 2 Then
                    imgSword.Top = PlayerImage(0).Top + PlayerImage(0).Height
                    imgSword.Left = PlayerImage(0).Left
                    imgSword.Picture = frmGraphics.imgSwordDown.Picture
                End If
                If Dir = 3 Then
                    imgSword.Top = PlayerImage(0).Top
                    imgSword.Left = PlayerImage(0).Left + PlayerImage(0).Width
                    imgSword.Picture = frmGraphics.imgSwordRight.Picture
                End If
                If Dir = 4 Then
                    imgSword.Top = PlayerImage(0).Top
                    imgSword.Left = PlayerImage(0).Left - PlayerImage(0).Width
                    imgSword.Picture = frmGraphics.imgSwordLeft.Picture
                End If
                imgSword.Visible = True

                tmrSword.Enabled = True
            End If
            
            If KeyCode = vbKeyEscape Then
                Mbox.Label1.Caption = "Exit?"
                Mbox.Label2.Caption = "Are you sure you want to exit Dragon Field?"
                Mbox.Show vbModal, Me
                If MBoxReturn = True Then
                    End
                Else
                    Unload Mbox
                End If
            End If
        End If ' end if player.direction.stopped = false
        If Player.Direction.Stopped = True Then
            If KeyCode = vbKeyF1 Then
                frmOptions.Show vbModal, Me
            End If
            
            If KeyCode = vbKeyA Then
                Scrolling = True
                PlayWav "Sword.wav"
                If Dir = 1 Then
                    imgSword.Top = PlayerImage(0).Top - PlayerImage(0).Height
                    imgSword.Left = PlayerImage(0).Left
                    imgSword.Picture = frmGraphics.imgSwordUp.Picture
                End If
                If Dir = 2 Then
                    imgSword.Top = PlayerImage(0).Top + PlayerImage(0).Height
                    imgSword.Left = PlayerImage(0).Left
                    imgSword.Picture = frmGraphics.imgSwordDown.Picture
                End If
                If Dir = 3 Then
                    imgSword.Top = PlayerImage(0).Top
                    imgSword.Left = PlayerImage(0).Left + PlayerImage(0).Width
                    imgSword.Picture = frmGraphics.imgSwordRight.Picture
                End If
                If Dir = 4 Then
                    imgSword.Top = PlayerImage(0).Top
                    imgSword.Left = PlayerImage(0).Left - PlayerImage(0).Width
                    imgSword.Picture = frmGraphics.imgSwordLeft.Picture
                End If
                imgSword.Visible = True
                DoEvents
                tmrSword.Enabled = True
            End If
            If KeyCode = vbKeyTab Then
                If frmMenu.Visible = False Then
                    frmMenu.Visible = True
                    Me.Height = Me.Height + (frmMenu.Height * 15.1)
                    GoTo DEF
                End If
                If frmMenu.Visible = True Then
                    Me.Height = Me.Height - (frmMenu.Height * 15.1)
                    frmMenu.Visible = False
                    GoTo DEF
                End If
DEF:
                If frmMenu.Visible = True And frmWeapons.Visible = True Then
                    imgMiniMap.Top = frmWeapons.Height
                    imgMiniMap.Left = frmMenu.Width
                    imgMiniMap.Height = (frmMainOnline.Height / Screen.TwipsPerPixelY) - frmWeapons.Height
                    imgMiniMap.Width = (frmMainOnline.Width / Screen.TwipsPerPixelX) - frmMenu.Width
                    imgMiniMap.Visible = True
                Else
                    imgMiniMap.Visible = False
                End If
                Exit Sub
            End If
            If KeyCode = vbKeyI Then
                frmIS.Show vbModal, Me
                Exit Sub
            End If
            If KeyCode = 192 Then
                If frmWeapons.Visible = False Then
                    frmWeapons.Visible = True
                    Me.Width = Me.Width + (frmWeapons.Width * 15.1)
                    GoTo ABC
                End If
                If frmWeapons.Visible = True Then
                    frmWeapons.Visible = False
                    Me.Width = Me.Width - (frmWeapons.Width * 15.1)
                    GoTo ABC
                End If
ABC:
                If frmMenu.Visible = True And frmWeapons.Visible = True Then
                    imgMiniMap.Top = frmWeapons.Height
                    imgMiniMap.Left = frmMenu.Width
                    imgMiniMap.Height = ((frmMainOnline.Height / Screen.TwipsPerPixelY) - frmWeapons.Height) - 25
                    imgMiniMap.Width = (frmMainOnline.Width / Screen.TwipsPerPixelX) - frmMenu.Width
                    imgMiniMap.Visible = True
                Else
                    imgMiniMap.Visible = False
                End If
                Exit Sub
        End If
        End If 'end if player.direction.stopped = true
            CheckWallPos

        Scrolling = True
        If Player.Map.NPCs.ScriptSource.Count = "0" Then GoTo End123
        For i = 1 To Player.Map.NPCs.ScriptSource.Count
            DoEvents
            LoadScriptOnline i
        Next i
End123:
        If Player.Map.Gold.Count = "0" Then GoTo End1234
        For i = 1 To Player.Map.Gold.Count
            DoEvents
            LoadGoldOnline i
        Next i
End1234:
        If Player.Map.Arrows.Count = "0" Then GoTo End12345
        For i = 1 To Player.Map.Arrows.Count
            DoEvents
            LoadArrowsOnline i
        Next i
End12345:
        If Player.Map.Bombs.Count = "0" Then GoTo End12346
        For i = 1 To Player.Map.Bombs.Count
            DoEvents
            LoadBombsOnline i
        Next i
End12346:
        If Player.Map.Bushes.Count = "0" Then GoTo End12347
        For i = 1 To Player.Map.Bushes.Count
            DoEvents
            LoadBushesOnline i
        Next i
End12347:

        Scrolling = False
End Sub

Private Sub txtCommands_KeyUp(KeyCode As Integer, Shift As Integer)
    PlayerStopped

End Sub

Public Function CheckForWalls(ByVal key As String) As Boolean
    Dim PX, PY As Integer
    Dim PX2, PY2 As Integer
    
    PX2 = (PlayerImage(0).Left / 32)
    PY2 = (PlayerImage(0).Top / 32)
    
    On Error GoTo skipit
    If Dir = 1 Then
        If Player.Map.Bushes((PX2 + 1) & "," & (PY2)) = "True" Then Exit Function
    End If
    If Dir = 2 Then
        If Player.Map.Bushes((PX2 + 1) & "," & (PY2 + 2)) = "True" Then Exit Function
    End If
    If Dir = 3 Then
        If Player.Map.Bushes((PX2 + 2) & "," & (PY2 + 1)) = "True" Then Exit Function
    End If
    If Dir = 4 Then
        If Player.Map.Bushes((PX2) & "," & (PY2 + 1)) = "True" Then Exit Function
    End If

skipit:
    If Right(PX2, 2) = ".5" Then
        If Dir = 3 Then
        PX = Mid(PX2, 1, Len(PX2) - 2)
        PX3 = PX
        End If
        If Dir = 4 Then
        PX = Mid(PX2, 1, Len(PX2) - 2) + 1
        PX3 = PX
        End If
        If Dir = 1 Then
            If Dir2 = 3 Then
            PX = Mid(PX2, 1, Len(PX2) - 2)
            End If
            If Dir2 = 4 Then
            PX = Mid(PX2, 1, Len(PX2) - 2) + 1
            End If
            If Dir = 1 Or Dir = 2 Then PX = PX3
        End If
        If Dir = 2 Then
            If Dir2 = 3 Then
            PX = Mid(PX2, 1, Len(PX2) - 2)
            End If
            If Dir2 = 4 Then
            PX = Mid(PX2, 1, Len(PX2) - 2) + 1
            End If
            If Dir = 1 Or Dir = 2 Then PX = PX3
        End If
    Else
        PX = PX2
    End If
    If Right(PY2, 2) = ".5" Then
        If Dir = 1 Then
        PY = Mid(PY2, 1, Len(PY2) - 2)
        PY3 = PY
        End If
        If Dir = 2 Then
        PY = Mid(PY2, 1, Len(PY2) - 2) + 1
        PY3 = PY
        End If
        If Dir = 3 Then
            If Dir2 = 1 Then
            PY = Mid(PY2, 1, Len(PY2) - 2)
            End If
            If Dir2 = 2 Then
            PY = Mid(PY2, 1, Len(PY2) - 2) + 1
            End If
            If Dir = 3 Or Dir = 4 Then PY = PY3
        End If
        If Dir = 4 Then
            If Dir2 = 1 Then
            PY = Mid(PY2, 1, Len(PY2) - 2)
            End If
            If Dir2 = 2 Then
            PY = Mid(PY2, 1, Len(PY2) - 2) + 1
            End If
            If Dir = 3 Or Dir = 4 Then PY = PY3
        End If
    Else
        PY = PY2
    End If

    
    Select Case key
    Case vbKeyUp
        If Player.Map.MapBoxes(PX & "," & PY - 1) = "unwalkable" Then
            CheckForWalls = False
            Exit Function
        Else
            CheckForWalls = True
            Exit Function
        End If
    Case vbKeyDown
        If Player.Map.MapBoxes(PX & "," & PY + 1) = "unwalkable" Then
            CheckForWalls = False
            Exit Function
        Else
            CheckForWalls = True
            Exit Function
        End If
    Case vbKeyLeft
        If Player.Map.MapBoxes(PX - 1 & "," & PY) = "unwalkable" Then
            CheckForWalls = False
            Exit Function
        Else
            CheckForWalls = True
            Exit Function
        End If
    Case vbKeyRight
        If Player.Map.MapBoxes(PX + 1 & "," & PY) = "unwalkable" Then
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

Public Function CheckForSigns() As Boolean
    On Error GoTo errore
    Dim ReadSignString As String
    Dim PlayerCords As String
        
    PlayerCords = Player.PlayerX + 1 & "," & Player.PlayerY
    ReadSignString = Player.Map.Signs.Item(PlayerCords)
    SignData = ReadSignString
    CheckForSigns = True
    Exit Function
errore:
    CheckForSigns = False
    Exit Function
End Function

Public Function CheckWallPos()
    On Error GoTo errored

    Dim PlayerCords As String
    Dim ReadWarpString As String
    Dim DwarpX As Integer
    Dim DwarpY As Integer
    Dim MapName As String
    Dim i As Integer
    
    
    PlayerCords = Player.PlayerX + 1 & "," & Player.PlayerY + 1

    ReadWarpString = Player.Map.Warps.Item(PlayerCords)
    If ReadWarpString = "" Then Exit Function

    i = 1
    Do
        If Mid(ReadWarpString, i, 1) = "," Then GoTo DoneDwarpX
        DwarpX = DwarpX & Mid(ReadWarpString, i, 1)
        i = i + 1
    Loop

DoneDwarpX:
    i = i + 1
    Do
        If Mid(ReadWarpString, i, 1) = "," Then GoTo DoneDwarpY
        DwarpY = DwarpY & Mid(ReadWarpString, i, 1)
        i = i + 1
    Loop

DoneDwarpY:
    i = i + 1
    Do Until i >= Len(ReadWarpString) + 1
        MapName = MapName & Mid(ReadWarpString, i, 1)
        i = i + 1
    Loop

    Warp DwarpX, DwarpY, MapName
    
    Exit Function
errored:

End Function

Public Function Warp(DwarpX As Integer, DwarpY As Integer, MapName As String)

    Dim i As Integer
    Dim TLine As Integer
    Dim Temp As String
    Dim Name, ScriptSource, GifPath As String
    Dim XPos, YPos, Wid, Hei As Integer
    Dim WarpX, WarpY As Integer
    Dim MapNum As Integer
    Dim t As Integer
    Dim ScrollSpeed As Integer
    Dim ST As String
    
    MapNum = FreeFile
    
    Open App.Path & "\Options.txt" For Input As #MapNum
        Input #MapNum, ScrollSpeed
    Close #MapNum
    
    MapNum = FreeFile
    Scrolling = True
    img.Picture = frmMainOnline.Image
    img3.Picture = frmMainOnline.Image
    
    
    
    
    Open App.Path & "\Maps\" & MapName For Input As MapNum
    
    Input #MapNum, Player.Map.MapWidth
    Input #MapNum, Player.Map.MapHeight
    Input #MapNum, Player.Map.MapName
    frmDB.txtLevel.Text = MapName
    Input #MapNum, Player.Map.Music
    
    Player.Map.TotalMapData = ""
    
    
    Do Until Temp = "*WARP*" Or EOF(MapNum)
    
        Line Input #MapNum, Temp
                
        Player.Map.MapData = ""
        If Player.Map.TotalMapData <> "" Then
            Player.Map.TotalMapData = Player.Map.TotalMapData & ","
        End If
        For i = 1 To Len(Temp)
        
            If IsNumeric(Mid(Temp, i, 1)) Or Mid(Temp, i, 1) = "," Then
                Player.Map.MapData = Player.Map.MapData + Mid(Temp, i, 1)
                Player.Map.TotalMapData = Player.Map.TotalMapData + Mid(Temp, i, 1)
            End If
            
        Next i
    
        DisplayMapLine TLine
    
        'Increment the Line Counter
        
        TLine = TLine + 1
    
    Loop
    
    If Player.Direction.Down = True Then
        DwarpX = DwarpX - 1
        DwarpY = DwarpY - 1
    End If
    If Player.Direction.Up = True Then
        DwarpY = DwarpY - 1
        DwarpX = DwarpX - 1
    End If
    If Player.Direction.Left = True Then
        DwarpX = DwarpX - 1
        DwarpY = DwarpY - 1
    End If
    If Player.Direction.Right = True Then
        DwarpY = DwarpY - 1
        DwarpX = DwarpX - 1
    End If
        
    DoEvents
    frmMainOnline.Refresh
    
    ModFormMapPos.Init
    Player.PlayerX = DwarpX
    Player.PlayerY = DwarpY
    PlayerImage(0).Left = DwarpX * 32
    PlayerImage(0).Top = DwarpY * 32
    PlayerImage(0).Refresh
    frmMusic.PlayMusic Player.Map.Music
    

    
    If EOF(MapNum) Then GoTo SkipWarp
    'Found the Warps
    
    i = 1: WarpX = "0": DwarpX = "0": WarpY = "0": DwarpY = "0": MapName = ""
        

    ClearWarps
    Do Until Temp = "*NPC*"
        Input #MapNum, WarpX
        If WarpX = "*NPC*" Then GoTo SkipWarp
        Input #MapNum, WarpY
        Input #MapNum, MapName
        Input #MapNum, DwarpX
        Input #MapNum, DwarpY
        Player.Map.Warps.Add DwarpX & "," & DwarpY & "," & MapName, WarpX & "," & WarpY
    Loop
    
SkipWarp:
    
    'NPCs
    ClearNPCs
   Do Until Temp = "*SIGNS*"
        ' NPC Format: Name, ScriptFile, GIF, X, Y, Height, Width
        Input #MapNum, Temp
        If Temp = "*SIGNS*" Then GoTo SkipNPCs
        Name = Temp
        Input #MapNum, Temp
        ScriptSource = Temp
        Input #MapNum, Temp
        GifPath = Temp
        Input #MapNum, Temp
        XPos = Temp
        Input #MapNum, Temp
        YPos = Temp
        Input #MapNum, Temp
        Hei = Temp
        Input #MapNum, Temp
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
        If Player.Map.NPCs.ScriptSource.Count = "0" Then GoTo SkipNPCs
        For i = 1 To Player.Map.NPCs.ScriptSource.Count
            DoEvents
            LoadScriptOnline i
        Next i
SkipNPCs:
    
    
    ClearSigns
    Temp = ""
    Do Until Temp = "*GOLD*"
        Line Input #MapNum, Temp
        If Temp = "*GOLD*" Then GoTo GGold
        i = 1
        Temp2 = ""
        Temp3 = ""
        Do
            If Mid(Temp, i, 1) = "," Or i >= Len(Temp) + 1 Then GoTo Next9
            Temp2 = Temp2 & Mid(Temp, i, 1)
            i = i + 1
        Loop
Next9:
i = i + 1
Temp2 = Temp2 & ","
        Do
            If Mid(Temp, i, 1) = "," Or i >= Len(Temp) + 1 Then GoTo Next10
            Temp2 = Temp2 & Mid(Temp, i, 1)
            i = i + 1
        Loop
Next10:
i = i + 1
        Do
            If i >= Len(Temp) + 1 Then GoTo Doneit
            Temp3 = Temp3 & Mid(Temp, i, 1)
            i = i + 1
        Loop
Doneit:

        
        Player.Map.Signs.Add Temp3, Temp2
    Loop
GGold:
        
Goldd:
    ClearGoldOnline
    Do Until Temp = "*ARROWS*"
        Line Input #MapNum, Temp
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
    ClearArrowsOnline
    Do Until Temp = "*BOMBS*"
        Line Input #MapNum, Temp
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
            imgArrow(Player.Map.Arrows.Count).Left = (Temp - 1) * 32
            imgArrow(Player.Map.Arrows.Count).Top = (Temp2 - 1) * 32
            imgArrow(Player.Map.Arrows.Count).Visible = True
        Close #24
        Kill App.Path & "\ArrowsTemp.txt"
    Loop
SkipArrows:
    ClearBombsOnline
    Do Until Temp = "*BUSHES*"
        Line Input #MapNum, Temp
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
            imgBomb(Player.Map.Bombs.Count).Left = (Temp - 1) * 32
            imgBomb(Player.Map.Bombs.Count).Top = (Temp2 - 1) * 32
            imgBomb(Player.Map.Bombs.Count).Visible = True
        Close #24
        Kill App.Path & "\BombsTemp.txt"
    Loop
SkipBombs:
    ClearBushesOnline
    Do Until EOF(MapNum)
        Line Input #MapNum, Temp
    Open App.Path & "\BushesTemp.txt" For Output As #24
            Print #24, Temp
        Close #24
        Open App.Path & "\BushesTemp.txt" For Input As #24
            Input #24, Temp
            Input #24, Temp2
            Player.Map.Bushes.Add "True", Temp & "," & Temp2
            'Bushes format: X,Y
            Load imgBush(Player.Map.Bushes.Count)
            imgBush(Player.Map.Bushes.Count).Picture = frmGraphics.imgBush.Picture
            imgBush(Player.Map.Bushes.Count).Left = (Temp - 1) * 32
            imgBush(Player.Map.Bushes.Count).Top = (Temp2 - 1) * 32
            imgBush(Player.Map.Bushes.Count).Visible = True
            Close #24
            Kill App.Path & "\BushesTemp.txt"
        Loop
    Close #MapNum
    Open App.Path & "\Options.txt" For Input As #MapNum
        Input #MapNum, ST
        Input #MapNum, ST
        Input #MapNum, ST
        Input #MapNum, ST
    Close #MapNum
    
    If ST = "No" Then GoTo skipst
    img2.Picture = frmDummy.Image
    img3.Picture = frmDummy.Image
    PlayerImage(0).Visible = False
        Do Until t >= frmMainOnline.ScaleHeight
            frmMainOnline.PaintPicture img2.Picture, 0, Me.ScaleHeight - t
            frmMainOnline.PaintPicture img.Picture, 0, t, Me.Width, Me.Height - t
            frmMainOnline.Refresh
            t = t + (ScrollSpeed * 32)
        Loop
    PlayerImage(0).Visible = True
    Me.Refresh
skipst:
    frmMainOnline.PaintPicture frmDummy.Image, 0, 0
    lblName.Caption = Player.Map.MapName
    lblName2.Caption = Player.Map.MapName
    lblName.Visible = True
    lblName2.Visible = True
    tmrName.Enabled = True
    Scrolling = False
    DoEvents
    LoadNPCs

End Function

Public Function PlayerWasHit(Power As Integer)
    Dim i As String
    Dim Evasion As Integer
    
    i = 0
    
    PlayWav "OOH1.wav"
    
    If Player.Evasion > "0" Then
        Randomize
        Evasion = Int((100 * Rnd) + 1)
        If Evasion < Player.Evasion Then Exit Function
    End If
    
    If Player.Defense > Power Then
        Power = Power * 0.15  'if defense is greater than power, hurt them 15% of the attack
    ElseIf Power > Player.Defense Then
        Power = Power - Player.Defense
    End If
    
    If Power >= txtHealth.Text Then GoTo PlayerIsDead
    
    If txtHealth.Text < 2 And Power <> Player.Defense Then GoTo PlayerIsDead
    
    tmrHit.Enabled = True
    txtHealth.Text = txtHealth - Power
    Player.Health = txtHealth.Text
    frmDB.txtHealth.Text = txtHealth.Text
    
    If Dir = 2 And Player.PlayerY >= 1 And Player.Map.MapBoxes(Player.PlayerX & "," & Player.PlayerY - 1) = "walkable" Then
        PlayerImage(0).Top = PlayerImage(0).Top - 32
        Player.PlayerY = Player.PlayerY - 1
    End If
    If Dir = 1 And Player.PlayerY <= 13 And Player.Map.MapBoxes(Player.PlayerX & "," & Player.PlayerY + 1) = "walkable" Then
        PlayerImage(0).Top = PlayerImage(0).Top + 32
        Player.PlayerY = Player.PlayerY + 1
    End If
    If Dir = 3 And Player.PlayerX >= 1 And Player.Map.MapBoxes(Player.PlayerX - 1 & "," & Player.PlayerY) = "walkable" Then
        PlayerImage(0).Left = PlayerImage(0).Left - 32
        Player.PlayerX = Player.PlayerX - 1
    End If
    If Dir = 4 And Player.PlayerX <= 20 And Player.Map.MapBoxes(Player.PlayerX + 1 & "," & Player.PlayerY) = "walkable" Then
        PlayerImage(0).Left = PlayerImage(0).Left + 32
        Player.PlayerX = Player.PlayerX + 1
    End If
    DoEvents
    Exit Function
    
PlayerIsDead:
    i = 0
    txtHealth.Text = "0"
    Player.Health = "0"
    Me.Refresh
    Me.Enabled = False
    Do Until i = 40000
        i = i + 1
        DoEvents
    Loop
    frmDB.txtDeaths.Text = frmDB.txtDeaths.Text + 1
    Me.txtHealth.Text = frmDB.txtTotalHealth.Text
    Player.Health = txtHealth.Text
    frmDB.txtHealth.Text = frmDB.txtTotalHealth.Text
    Player.Direction.Dead = True
    DoEvents
    Unload Me
    frmDied.Show
    Exit Function
End Function

Private Sub txtHealth_Change()
    txtHealthDSP.Text = txtHealth.Text & "/" & txtTotalHealth.Text
End Sub

Private Sub txtTotalHealth_Change()
    txtHealthDSP.Text = txtHealth.Text & "/" & txtTotalHealth.Text
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

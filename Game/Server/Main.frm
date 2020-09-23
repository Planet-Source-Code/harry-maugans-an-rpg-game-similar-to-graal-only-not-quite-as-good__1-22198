VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Main 
   Caption         =   "Server"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11400
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10590
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAccess 
      Height          =   285
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   61
      Text            =   "txtAccess"
      Top             =   8400
      Width           =   1935
   End
   Begin VB.TextBox txtBanned 
      Height          =   285
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   60
      Text            =   "txtBanned"
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox txtPlayerY 
      Height          =   285
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   59
      Text            =   "txtPlayerY"
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox txtPlayerX 
      Height          =   285
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   58
      Text            =   "txtPlayerX"
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox txtLevel 
      Height          =   285
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   57
      Text            =   "txtLevel"
      Top             =   6960
      Width           =   1935
   End
   Begin VB.TextBox txtHealth 
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   56
      Text            =   "txtHealth"
      Top             =   8400
      Width           =   1935
   End
   Begin VB.TextBox txtGold 
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   "txtGold"
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox txtDeaths 
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   54
      Text            =   "txtDeaths"
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox txtKills 
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   53
      Text            =   "txtKills"
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox txtIdleTime 
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "txtIdleTime"
      Top             =   6960
      Width           =   1935
   End
   Begin VB.TextBox txtLastSignon 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   51
      Text            =   "txtLastSignon"
      Top             =   8400
      Width           =   1935
   End
   Begin VB.TextBox txtOnlineTime 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "txtOnlineTime"
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "txtID"
      Top             =   6960
      Width           =   1935
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "txtPassword"
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "txtUsername"
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox txtTotalHealth 
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "txtTotalHealth"
      Top             =   8760
      Width           =   1935
   End
   Begin VB.TextBox txtWeaponCount 
      Height          =   285
      Left            =   3240
      TabIndex        =   45
      Text            =   "txtWeaponCount"
      Top             =   8760
      Width           =   1935
   End
   Begin VB.ListBox lstWeapons 
      Height          =   645
      Left            =   3240
      TabIndex        =   44
      Top             =   9120
      Width           =   1935
   End
   Begin VB.TextBox txtActiveWeapon 
      Height          =   285
      Left            =   3240
      TabIndex        =   43
      Text            =   "txtActiveWeapon"
      Top             =   9840
      Width           =   1935
   End
   Begin VB.TextBox txtArmor 
      Height          =   285
      Left            =   6360
      TabIndex        =   42
      Text            =   "txtArmor"
      Top             =   9240
      Width           =   1935
   End
   Begin VB.TextBox txtShield 
      Height          =   285
      Left            =   6360
      TabIndex        =   41
      Text            =   "txtShield"
      Top             =   9600
      Width           =   1935
   End
   Begin VB.TextBox txtSword 
      Height          =   285
      Left            =   6360
      TabIndex        =   40
      Text            =   "txtSword"
      Top             =   9960
      Width           =   1935
   End
   Begin VB.TextBox txtHelmet 
      Height          =   285
      Left            =   3240
      TabIndex        =   39
      Text            =   "txtHelmet"
      Top             =   10200
      Width           =   1935
   End
   Begin VB.ListBox lstInventory 
      Height          =   645
      Left            =   9480
      TabIndex        =   38
      Top             =   9600
      Width           =   1935
   End
   Begin VB.TextBox txtArrows 
      Height          =   285
      Left            =   9480
      TabIndex        =   37
      Text            =   "txtArrows"
      Top             =   8880
      Width           =   1935
   End
   Begin VB.TextBox txtBombs 
      Height          =   285
      Left            =   9480
      TabIndex        =   36
      Text            =   "txtBombs"
      Top             =   9240
      Width           =   1935
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Test Out Level"
      Height          =   285
      Left            =   6960
      TabIndex        =   32
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Shell Map Editor"
      Height          =   285
      Left            =   5400
      TabIndex        =   31
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Edit Account Status"
      Height          =   255
      Left            =   3480
      TabIndex        =   27
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Disable Account"
      Height          =   255
      Left            =   2040
      TabIndex        =   29
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Delete Account"
      Height          =   255
      Left            =   3480
      TabIndex        =   28
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Upload/Update A Level"
      Height          =   255
      Left            =   6960
      TabIndex        =   13
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Reset Levels"
      Height          =   255
      Left            =   5400
      TabIndex        =   12
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selected User's Information:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   1920
      TabIndex        =   11
      Top             =   3840
      Width           =   3255
      Begin VB.Label Label6 
         Caption         =   "Account Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Account Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Account Status:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Remote Server Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Remote IP Address:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lblSocketNumber 
         Caption         =   "Socket Number:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Edit Account"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Shut Down Server"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reset Server"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update User Info"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ban User"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Boot User"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.ListBox lstMembers 
      Height          =   3960
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   8760
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Local Information:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5400
      TabIndex        =   1
      Top             =   3840
      Width           =   3855
      Begin VB.Label lblRCs 
         Caption         =   "Current Number of RCs Connected:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label lblLocalServer 
         Caption         =   "Local Server Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label lblSocketsConnected 
         Caption         =   "Number of Sockets Connected:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label lblMaxSockets 
         Caption         =   "Max Number of Sockets: unlimited"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label lblLocalPort 
         Caption         =   "Local Port Listening To:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label lblLocalIP 
         Caption         =   "Local IP Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lblLevels 
         Caption         =   "Current Total Number of Levels:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   3615
      End
   End
   Begin VB.ListBox Status 
      Appearance      =   0  'Flat
      Height          =   3345
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   9375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Edit News"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Suspend Account"
      Height          =   255
      Left            =   2040
      TabIndex        =   30
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label34 
      Caption         =   "Access Priviledges:"
      Height          =   495
      Left            =   8520
      TabIndex        =   87
      Top             =   8385
      Width           =   855
   End
   Begin VB.Label Label33 
      Caption         =   "Banned:"
      Height          =   255
      Left            =   8520
      TabIndex        =   86
      Top             =   8040
      Width           =   855
   End
   Begin VB.Label Label32 
      Caption         =   "Player Y:"
      Height          =   255
      Left            =   8520
      TabIndex        =   85
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label31 
      Caption         =   "Player X:"
      Height          =   255
      Left            =   8520
      TabIndex        =   84
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label Label30 
      Caption         =   "Level:"
      Height          =   255
      Left            =   8520
      TabIndex        =   83
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label29 
      Caption         =   "Health:"
      Height          =   255
      Left            =   5400
      TabIndex        =   82
      Top             =   8400
      Width           =   855
   End
   Begin VB.Label Label28 
      Caption         =   "Gold:"
      Height          =   255
      Left            =   5400
      TabIndex        =   81
      Top             =   8040
      Width           =   855
   End
   Begin VB.Label Label27 
      Caption         =   "Deaths:"
      Height          =   255
      Left            =   5400
      TabIndex        =   80
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label21 
      Caption         =   "Kills:"
      Height          =   255
      Left            =   5400
      TabIndex        =   79
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label Label20 
      Caption         =   "Idle Time:"
      Height          =   255
      Left            =   5400
      TabIndex        =   78
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label19 
      Caption         =   "Last Signon:"
      Height          =   255
      Left            =   2040
      TabIndex        =   77
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label Label18 
      Caption         =   "Online Time:"
      Height          =   255
      Left            =   2040
      TabIndex        =   76
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "Member ID:"
      Height          =   255
      Left            =   2040
      TabIndex        =   75
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "Password:"
      Height          =   255
      Left            =   2040
      TabIndex        =   74
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "Username:"
      Height          =   255
      Left            =   2040
      TabIndex        =   73
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Total Health:"
      Height          =   375
      Left            =   5400
      TabIndex        =   72
      Top             =   8760
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Weapon Count:"
      Height          =   255
      Left            =   2040
      TabIndex        =   71
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Weapons:"
      Height          =   255
      Left            =   2040
      TabIndex        =   70
      Top             =   9120
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Active Weapon:"
      Height          =   255
      Left            =   2040
      TabIndex        =   69
      Top             =   9840
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Armor:"
      Height          =   255
      Left            =   5400
      TabIndex        =   68
      Top             =   9240
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Shield:"
      Height          =   255
      Left            =   5400
      TabIndex        =   67
      Top             =   9600
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Sword:"
      Height          =   255
      Left            =   5400
      TabIndex        =   66
      Top             =   9960
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Helmet:"
      Height          =   255
      Left            =   2400
      TabIndex        =   65
      Top             =   10200
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "Inventory:"
      Height          =   255
      Left            =   8400
      TabIndex        =   64
      Top             =   9600
      Width           =   1095
   End
   Begin VB.Label Label25 
      Caption         =   "Arrows:"
      Height          =   255
      Left            =   8520
      TabIndex        =   63
      Top             =   8880
      Width           =   855
   End
   Begin VB.Label Label26 
      Caption         =   "Bombs:"
      Height          =   255
      Left            =   8520
      TabIndex        =   62
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label lblMembersOnline 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   35
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label23 
      Caption         =   "Changes when checking for a password or loading someone's stats, or anytime something is changed."
      Height          =   1215
      Left            =   0
      TabIndex        =   34
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label22 
      Caption         =   "System Information:"
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
      Left            =   0
      TabIndex        =   33
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11400
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Image Image1 
      Height          =   2280
      Left            =   9360
      Picture         =   "Main.frx":0442
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "Members Online:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iSockets As Integer
Dim sRequestID As String

Private Sub Form_Load()
    
    lblLocalServer.Caption = Sock(0).LocalHostName
    lblLocalIP.Caption = Sock(0).LocalIP
    Sock(0).LocalPort = "6005"
    Status.AddItem "Server started at: " & Time & " on " & Date
    Status.AddItem "Listening to port: " & Sock(0).LocalPort
    Sock(0).Listen
    
    DatabaseStartup
    LoadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DB.Close
End Sub

Private Sub Sock_Close(Index As Integer)
    Status.AddItem "Connection Closed: " & Sock(Index).RemoteHostIP & " at " & Time & " on " & Date
    Sock(Index).Close
    Unload Sock(Index)
    iSockets = iSockets - 1
    lblSocketsConnected.Caption = "Number of Sockets Connected: " & iSockets
    lblMembersOnline.Caption = iSockets
End Sub

Private Sub Sock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If Index = 0 Then
        Status.AddItem "Connection request from: " & Sock(Index).RemoteHostIP & " at " & Time & " on " & Date
        sRequestID = requestID
        iSockets = iSockets + 1
        lblSocketsConnected.Caption = "Number of Sockets Connected: " & iSockets
        lblMembersOnline.Caption = iSockets
        Load Sock(iSockets)
        Sock(iSockets).LocalPort = "6005"
        Sock(iSockets).Accept requestID
    End If
End Sub

Private Sub Sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim InData  As String
    Dim OutData As String
    Dim Temp As String
    Dim Username As String
    Dim Password As String
    Dim i As Integer
    
    Sock(Index).GetData InData, vbString, 10000  ' max length string accepted is 10 megs
    
    Select Case Mid(InData, 1, 11)
        Case "clientlogin":
            InData = Mid(InData, 13, Len(InData) - 10)
            i = 1
            Username = ""
            Do Until i > Len(InData)
                Temp = Mid(InData, i, 1)
                If Temp = "," Then GoTo FoundComma
                Username = Username & Temp
                i = i + 1
            Loop
FoundComma:
            Password = ""
            i = i + 1
            Do Until i > Len(InData)
                Temp = Mid(InData, i, 1)
                Password = Password & Temp
                i = i + 1
            Loop
            
            'have username and password as variables, now check them
            OutData = CheckPassword(Username, Password)
            Sock(Index).SendData OutData
        Case "MakeAccount":
            'MakeAccount Data - goes - here
        Case "PlayerMoved":
            For i = 1 To iSockets
                Sock(i).SendData InData
            Next i
        Case Else:
            Sock(Index).SendData "Improper format of data."
            Sock_Close Index
    End Select
End Sub


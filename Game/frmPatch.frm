VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmPatch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Patch Wizard 1.2"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtHidden 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"frmPatch.frx":0000
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1320
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://zack_knall:30359@"
      UserName        =   "zack_knall"
      Password        =   "30359"
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "frmPatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Allow As Boolean
Dim Data As String
Dim Version As String
Dim Date2 As String
Dim Size As String
Dim Date3 As String

'Public oUnZip As New CGUnzipFiles
    
Private Sub Form_Load()
    List1.Clear
    Me.Visible = True
    List1.AddItem "Initializing Patch Program..."
    DoEvents
    Me.Refresh
    List1.Refresh
    List1.AddItem "Connecting to patch server..."
    DoEvents
    Me.Refresh
    List1.Refresh
    Data = Inet1.OpenURL("http://www.geocities.com/zack_knall/CurrentVersion.txt")
    DoEvents
    Me.Refresh
    List1.Refresh
    List1.AddItem "Checking Current Version..."
    DoEvents
    Me.Refresh
    List1.Refresh
    Open App.Path & "\OnlineTemp.txt" For Output As #1
        Print #1, Data
    Close #1
    Open App.Path & "\OnlineTemp.txt" For Input As #1
        Input #1, Version
        Input #1, Date2
    Close #1
    Kill App.Path & "\OnlineTemp.txt"
    List1.AddItem "Checking Your Version..."
    DoEvents
    Me.Refresh
    List1.Refresh
    Open App.Path & "\Options.txt" For Input As #1
        Line Input #1, Data
        Line Input #1, Data
        Line Input #1, Data
        Line Input #1, Data
        Line Input #1, Data
    Close #1
    DoEvents
    Me.Refresh
    List1.Refresh
    If Data >= Version Then
        List1.AddItem "Your version is the most current that is availible."
        DoEvents
        Me.Refresh
        List1.Refresh
        Allow = True
        Unload Me
    Else
        List1.AddItem "Your version is old.  We will now upgrade it."
        DoEvents
        Me.Refresh
        List1.Refresh
        UpgradeVersion
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If Allow = False Then
        Cancel = True
    'Else
    '    frmMainOnline.Show
    'End If
End Sub

Public Function UpgradeVersion()
    Dim i As Integer
    
    Data = Inet1.OpenURL("http://www.geocities.com/zack_knall/UpgradeSize.txt")
    
    Open App.Path & "\OnlineTemp.txt" For Output As #1
        Print #1, Data
    Close #1
    Open App.Path & "\OnlineTemp.txt" For Input As #1
        Input #1, Size
        Input #1, Date3
    Close #1
    Kill App.Path & "\OnlineTemp.txt"
    
    List1.AddItem "The newest release is version " & Version & "!"
    DoEvents
    Me.Refresh
    List1.Refresh
    List1.AddItem "It was last updated on " & Date2 & "."
    DoEvents
    Me.Refresh
    List1.Refresh
    List1.AddItem "This may take a few minutes, depending on your connection speed."
    DoEvents
    Me.Refresh
    List1.Refresh
    List1.AddItem "Downloading...Please wait"
    DoEvents
    Me.Refresh
    List1.Refresh
    Data = Inet1.OpenURL("http://www.geocities.com/zack_knall/GameEXE.zip")
    List1.AddItem "Done downloading..."
    List1.AddItem "Loading file from memory to temp. buffer..."
    DoEvents
    Me.Refresh
    List1.Refresh
    txtHidden.Text = Data
    List1.AddItem "Saving file from temp buffer to file..."
    DoEvents
    Me.Refresh
    List1.Refresh
    txtHidden.SaveFile App.Path & "\Update-" & Version & "-" & Date3 & ".zip"
    List1.AddItem "File saved to the game's directory as a zip file."
    List1.AddItem "The zip file is: " & "Update-" & Version & "-" & Date3 & ".zip"
    List1.AddItem "We shall now unzip it."
    DoEvents
    Me.Refresh
    List1.Refresh
    With oUnZip
        .ZipFileName = App.Path & "\Update-" & Version & "-" & Date3 & ".zip"
        .ExtractDir = App.Path & "\..\"
        .HonorDirectories = True
    End With
    List1.AddItem "Deleting downloaded zip file..."
    Kill App.Path & "\Update-" & Version & "-" & Date3 & ".zip"
    List1.AddItem "Downloaded ZIP file has been deleted."
    DoEvents
    Me.Refresh
    List1.Refresh
    List1.AddItem "Game has been updated.  Please wait while your game restarts."
    DoEvents
    Me.Refresh
    List1.Refresh
    Do Until i = 32000
        i = i + 1
    Loop
    Shell App.Path & "\Game.exe"
    End
End Function

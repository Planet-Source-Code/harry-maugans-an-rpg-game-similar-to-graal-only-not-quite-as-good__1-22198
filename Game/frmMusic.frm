VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmMusic 
   BorderStyle     =   0  'None
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   661
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "Sequencer"
      FileName        =   "C:\My Documents\My Programs\Game\Music\wildarms.mid"
   End
End
Attribute VB_Name = "frmMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Function PlayMusic(FileName As String)
    On Error Resume Next
    Dim Music As String
    
    MMControl1.Command = "Stop"
    MMControl1.Command = "Close"
    
    Open App.Path & "\Options.txt" For Input As #9
        Input #9, Music
        Input #9, Music
        Input #9, Music
    Close #9
    
    If Music = "Off" Then Exit Function
    If frmMain.SoundCardInstalled = False Then MsgBox "No sound card detected!", vbCritical: Exit Function

    MMControl1.FileName = App.Path & "\Music\" & FileName
    MMControl1.Command = "Open"
    MMControl1.Command = "Play"

End Function


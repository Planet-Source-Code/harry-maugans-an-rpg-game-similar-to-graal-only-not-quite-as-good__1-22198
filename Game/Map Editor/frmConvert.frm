VERSION 5.00
Begin VB.Form frmConvert 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Map Converter"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   Icon            =   "frmConvert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   3240
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert!"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   4455
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "Whatever.map"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtOld 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "Whatever.map"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "New Map Name:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Old Map Name:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   $"frmConvert.frx":000C
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error GoTo errord
    Dim Temp As String
    
    Open App.Path & "\..\Maps\" & txtOld.Text For Input As #1
    Open App.Path & "\..\Maps\Temp1A2B3CSys.map" For Output As #2
        Do Until EOF(1)
            Line Input #1, Temp
            Print #2, Temp
        Loop
    Close #1
    Close #2
    

    Open App.Path & "\..\Maps\" & txtNew.Text For Output As #1
    Open App.Path & "\..\Maps\Temp1A2B3CSys.map" For Input As #2
        Do Until EOF(2)
            Line Input #2, Temp
            Print #1, Temp
        Loop
    Close #1
    Close #2
    Kill App.Path & "\..\Maps\Temp1A2B3CSys.map"
    
    MsgBox "Converted Successfully.", , "Done."
    Unload Me
    Exit Sub
errord:
    MsgBox Err.Description
    MsgBox "Cannot find file or unable to read properly.", , "Error"
    Unload Me
End Sub

Private Sub File1_Click()
    txtOld.Text = File1.FileName
    txtNew.Text = File1.FileName
End Sub

Private Sub Form_Load()
    File1.Path = App.Path & "\..\Maps\"
End Sub

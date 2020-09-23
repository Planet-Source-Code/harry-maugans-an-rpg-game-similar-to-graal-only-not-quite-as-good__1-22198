VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Load File"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Map"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   3855
   End
   Begin VB.FileListBox File1 
      Height          =   5160
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Please choose a map file to load:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmMain.DrawMap File1.FileName
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim Result As String
    If File1.FileName <> "" Then
        Result = MsgBox("Are you want to delete " & File1.FileName & " ?", vbYesNo, "Sure?")
        If Result = vbYes Then
            Kill (App.Path & "\..\Maps\" & File1.FileName)
        End If
    End If
    Command3_Click
End Sub

Private Sub Command3_Click()
    File1.Refresh
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub File1_DblClick()
    Command1_Click
End Sub

Private Sub Form_Load()
    File1.Path = App.Path & "\..\Maps\"
End Sub

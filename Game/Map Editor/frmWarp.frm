VERSION 5.00
Begin VB.Form frmWarp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Warp Editor"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDY 
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtDX 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtMap 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4440
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label5 
      Caption         =   "Destination Warp Y:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Destination Warp X:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Destination Map Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Warp Y:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Warp X:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmWarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Editing As Boolean

Private Sub Command1_Click()
If Editing = False Then
    txtX.Locked = True
    txtY.Locked = True
    If Trim(txtMap.Text) = "" Or Trim(txtDX.Text) = "" Or Trim(txtDY.Text) = "" Then
        MsgBox "Please enter all the need info to make the warp.", , "Error"
        Exit Sub
    End If
    frmMain.List1.AddItem txtX.Text & "," & txtY.Text & "," & txtMap.Text & "," & txtDX.Text & "," & txtDY.Text
    Unload Me
End If
If Editing = True Then
    txtX.Locked = False
    txtY.Locked = False
    frmMain.List1.RemoveItem frmMain.List1.ListIndex
    If Trim(txtMap.Text) = "" Or Trim(txtDX.Text) = "" Or Trim(txtDY.Text) = "" Then
        MsgBox "Please enter all the need info to make the warp.", , "Error"
        Exit Sub
    End If
    frmMain.List1.AddItem txtX.Text & "," & txtY.Text & "," & txtMap.Text & "," & txtDX.Text & "," & txtDY.Text
    Unload Me
End If
End Sub

Private Sub Command2_Click()
    Editing = False
    Unload Me
End Sub

Private Sub Form_Load()
    txtX.Text = (frmMain.Shape3.Left / 32) + 1
    txtY.Text = (frmMain.Shape3.Top / 32) + 1
    If Editing = True Then
        txtX.Locked = False
        txtY.Locked = False
    ElseIf Editing = False Then
        txtX.Locked = True
        txtY.Locked = True
    End If
End Sub


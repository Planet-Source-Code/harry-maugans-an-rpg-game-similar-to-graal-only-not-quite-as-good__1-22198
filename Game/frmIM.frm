VERSION 5.00
Begin VB.Form frmIM 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmIM.frx":0000
   ScaleHeight     =   945
   ScaleWidth      =   1920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMsg 
      Height          =   285
      Left            =   18000
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   150
   End
   Begin VB.TextBox txtMessage 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1680
   End
End
Attribute VB_Name = "frmIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Const RGN_COPY = 5
Private Const CreatedBy = "VBSFC 6.2"
Private Const RegisteredTo = "Not Registered"
Private ResultRegion As Long
Private Function CreateFormRegion(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long, Counter As Integer
    Dim PolyPoints() As POINTAPI
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)
    ObjectRegion = CreateRectRgn(148 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 90 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 396 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 213 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(123 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 43 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 470 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 274 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 4)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(218 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 45 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 265 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 115 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    ObjectRegion = CreateRectRgn(168 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 20 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 407 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 195 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 4)
    DeleteObject ObjectRegion
    ReDim PolyPoints(0 To 22)
    For Counter = 0 To 22
        PolyPoints(Counter).x = GP0X(Counter) * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX
        PolyPoints(Counter).y = GP0Y(Counter) * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 23, 1)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function
Private Function GP0X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0X = 12
    Case 1
        GP0X = 117
    Case 2
        GP0X = 118
    Case 3
        GP0X = 119
    Case 4
        GP0X = 122
    Case 5
        GP0X = 123
    Case 6
        GP0X = 123
    Case 7
        GP0X = 122
    Case 8
        GP0X = 122
    Case 9
        GP0X = 118
    Case 10
        GP0X = 115
    Case 11
        GP0X = 11
    Case 12
        GP0X = 10
    Case 13
        GP0X = 7
    Case 14
        GP0X = 2
    Case 15
        GP0X = 1
    Case 16
        GP0X = 0
    Case 17
        GP0X = 1
    Case 18
        GP0X = 6
    Case 19
        GP0X = 6
    Case 20
        GP0X = 7
    Case 21
        GP0X = 7
    Case 22
        GP0X = 10
    End Select
End Function
Private Function GP0Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0Y = 5
    Case 1
        GP0Y = 5
    Case 2
        GP0Y = 6
    Case 3
        GP0Y = 6
    Case 4
        GP0Y = 9
    Case 5
        GP0Y = 12
    Case 6
        GP0Y = 52
    Case 7
        GP0Y = 53
    Case 8
        GP0Y = 54
    Case 9
        GP0Y = 58
    Case 10
        GP0Y = 59
    Case 11
        GP0Y = 59
    Case 12
        GP0Y = 60
    Case 13
        GP0Y = 61
    Case 14
        GP0Y = 63
    Case 15
        GP0Y = 64
    Case 16
        GP0Y = 63
    Case 17
        GP0Y = 62
    Case 18
        GP0Y = 51
    Case 19
        GP0Y = 11
    Case 20
        GP0Y = 10
    Case 21
        GP0Y = 9
    Case 22
        GP0Y = 6
    End Select
End Function

Private Sub Form_Load()
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hWnd, CreateFormRegion(1, 1, 0, 0), True)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
End Sub

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
Me.Hide
frmMain.SetFocus
End Sub

Private Sub txtMessage_LostFocus()
Me.Hide
frmMain.SetFocus
End Sub


Private Sub txtMsg_KeyDown(KeyCode As Integer, Shift As Integer)
Me.Hide
frmMain.SetFocus
End Sub

Private Sub txtMsg_LostFocus()
Me.Hide
frmMain.SetFocus
End Sub


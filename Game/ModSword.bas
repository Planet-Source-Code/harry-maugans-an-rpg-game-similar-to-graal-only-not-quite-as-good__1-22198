Attribute VB_Name = "ModSword"

Option Explicit

' BitBlt API
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Const SRCCOPY = &HCC0020


Dim xSprite As Integer  ' the x of the sprite
Dim SwordSwinging As Boolean

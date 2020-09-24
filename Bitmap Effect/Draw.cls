VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Draws"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function PolyBezier Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long



Private Type POINTAPI
        x As Long
        y As Long
End Type
Function DrawCircle(x, y, HalfRad, Color As Long)

Dim XX, YY
Dim A As Single
Dim Ang As Double

Do Until A = 359
 
 Ang = Rad * A
 
 XX = x + (Sin(Ang) * HalfRad)
 YY = y + (Cos(Ang) * HalfRad)
 
 Call DrawPix(XX, YY, Color, True)
 'SetPixel Form1.PD.hdc, XX, YY, vbRed
 
 A = A + 1
 
Loop

'SetDIBits ODC, OBmp, 0, H, Pix(1, 1, 1), BI, 0&

End Function


Function DrawPix(x, y, Color As Long, AntiAlias As Boolean)

Dim Col As ExRGB
On Error Resume Next
Col = ExtractRGB(Color)

Pix(3, x, y) = Col.R
If Pix(3, x + 1, y) <> Col.R Then Pix(3, x + 1, y) = (Pix(3, x + 1, y) + (Col.R * 0.25)) / 2
If Pix(3, x, y + 1) <> Col.R Then Pix(3, x, y + 1) = (Pix(3, x, y + 1) + (Col.R * 0.25)) / 2
If Pix(3, x - 1, y) <> Col.R Then Pix(3, x - 1, y) = (Pix(3, x - 1, y) + (Col.R * 0.25)) / 2
If Pix(3, x, y - 1) <> Col.R Then Pix(3, x, y - 1) = (Pix(3, x, y - 1) + (Col.R * 0.25)) / 2

Pix(2, x, y) = Col.G
If Pix(2, x + 1, y) <> Col.G Then Pix(2, x + 1, y) = (Pix(2, x + 1, y) + (Col.G * 0.25)) / 2
If Pix(2, x, y + 1) <> Col.G Then Pix(2, x, y + 1) = (Pix(2, x, y + 1) + (Col.G * 0.25)) / 2
If Pix(2, x - 1, y) <> Col.G Then Pix(2, x - 1, y) = (Pix(2, x - 1, y) + (Col.G * 0.25)) / 2
If Pix(2, x, y - 1) <> Col.G Then Pix(2, x, y - 1) = (Pix(2, x, y - 1) + (Col.G * 0.25)) / 2

Pix(1, x, y) = Col.B
If Pix(1, x + 1, y) <> Col.B Then Pix(1, x + 1, y) = (Pix(1, x + 1, y) + (Col.B * 0.25)) / 2
If Pix(1, x, y + 1) <> Col.B Then Pix(1, x, y + 1) = (Pix(1, x, y + 1) + (Col.B * 0.25)) / 2
If Pix(1, x - 1, y) <> Col.B Then Pix(1, x - 1, y) = (Pix(1, x - 1, y) + (Col.B * 0.25)) / 2
If Pix(1, x, y - 1) <> Col.B Then Pix(1, x, y - 1) = (Pix(1, x, y - 1) + (Col.B * 0.25)) / 2


SetDIBits ODC, OBmp, 0, H, Pix(1, 1, 1), BI, 0&

End Function


Private Function ExtractRGB(Color As Long) As ExRGB

With ExtractRGB
 
  .R = Color Mod 256
  .G = (Color / 256) Mod 256
  .B = (Color / 65536)
 
End With

End Function

Function WrawText(Text As String, x, y, Color As Long, Transparent As Boolean) As Boolean

If Transparent = True Then
 SetBkMode ODC, 1
Else
 SetBkMode ODC, 0
End If

SetTextColor ODC, Color

TextOut ODC, x, y, Text, Len(Text)

End Function

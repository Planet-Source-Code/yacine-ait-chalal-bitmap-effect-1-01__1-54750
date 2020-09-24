VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Edtion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long

Private Type ObjectInfo

 NombreObject As Integer
 WD() As Integer
 HG() As Integer
 Pixs() As Byte
  
End Type

Public Sub CopyImage()

  Dim NewHDC  As Long
  Dim ClipBMP As Long
  Dim OldBMP  As Long
  
    NewHDC = CreateCompatibleDC(ODC)
  If (NewHDC = 0) Then Exit Sub
  ClipBMP = CreateCompatibleBitmap(ODC, W, H)
  If (ClipBMP = 0) Then Exit Sub
  OldBMP = SelectObject(NewHDC, ClipBMP)
  
  Call BitBlt(NewHDC, 0, 0, W, H, ODC, 0, 0, &HCC0020)
  Call SelectObject(NewHDC, OldBMP)
  
  Call EmptyClipboard
  Call OpenClipboard(0&)
  Call SetClipboardData(2, ClipBMP)
  Call CloseClipboard
  'close
  Call DeleteObject(NewHDC)
  
  MsgBox "Ok"
  
End Sub

Function LoadObject(FName As String, NombreObject As Integer)
Dim OI As ObjectInfo
Dim XX, YY
Dim R As Integer, G As Integer, B As Integer

Open FName For Binary As 2

 Get 2, 1, OI

Close 2



For XX = 1 To OI.WD(NombreObject) - 1
 For YY = 1 To OI.HG(NombreObject) - 1
  
   Pix(3, XX, YY) = OI.Pixs(NombreObject, 3, XX, YY)
   Pix(2, XX, YY) = OI.Pixs(NombreObject, 2, XX, YY)
   Pix(1, XX, YY) = OI.Pixs(NombreObject, 1, XX, YY)
  
  Next
 Next

SetDIBits ODC, OBmp, 0, H, Pix(1, 1, 1), BI, 0&

End Function
Function addObject(FName As String)

Dim OI As ObjectInfo
Dim Max As Integer
Dim XX, YY
Dim R, G, B

On Error Resume Next
Open FName For Binary As 1

 Get 1, 1, OI

Close 1

Max = OI.NombreObject + 1

ReDim Preserve OI.HG(1 To Max)
ReDim Preserve OI.WD(1 To Max)
'ReDim Preserve OI.Blue(1 To W, 1 To H, 1 To Max)
'ReDim Preserve OI.Green(1 To W, 1 To H, 1 To Max)
'ReDim Preserve OI.Red(1 To W, 1 To H, 1 To Max)
ReDim Preserve OI.Pixs(1 To Max, 1 To 4, 1 To W, 1 To H)

OI.NombreObject = Max
OI.WD(Max) = W
OI.HG(Max) = H

 For YY = 1 To H
  For XX = 1 To W
  
   OI.Pixs(Max, 3, XX, YY) = Pix(3, XX, YY)
   OI.Pixs(Max, 2, XX, YY) = Pix(2, XX, YY)
   OI.Pixs(Max, 1, XX, YY) = Pix(1, XX, YY)

 
  'OI.Pixs(Max, 3, XX, YY) = (R)
  'OI.Pixs(Max, 2, XX, YY) = (G)
  'OI.Pixs(Max, 1, XX, YY) = (B)
   
 Next
Next

Open FName For Binary As 1

 Put 1, 1, OI

Close 1



End Function
Function Loading(hdc As Long, Hbmp As Long, WD, HG) As Boolean

W = WD
H = HG

ReDim Pix(1 To 4, 1 To W, 1 To H)

 With BI.Header
 
    .biSize = 40
    .biWidth = W
    .biHeight = -H
    .biPlanes = 1
    .biBitCount = 32
    .biSizeImage = 3 * W * H
    
 End With
 
 GetDIBits hdc, Hbmp, 0, H, Pix(1, 1, 1), BI, 0&
 
End Function


Function Clear()

ReDim Pix(1 To 4, 1 To W, 1 To H)

End Function

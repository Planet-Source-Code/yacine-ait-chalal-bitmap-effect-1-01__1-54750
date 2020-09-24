VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "E_Arts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum T_ArtLine

   AL_Vertical = 1
   AL_Horizontal = 2

End Enum
Const AlphaStepVal = (1 / 255)
Function ArtCircle(CX, CY, Color As Long)

Dim RStep
Dim A As Double
Dim Col As ExRGB
Dim R, G, B

On Error Resume Next

Col = ExtractRGB(Color)


For RStep = 2 To W Step 6
A = 0

 Do Until A > 6.28318

  XX = CX + Sin(A) * RStep
  YY = CY + Cos(A) * RStep
  A = A + Rad
 
  R = Pix(3, XX, YY)
  G = Pix(2, XX, YY)
  B = Pix(1, XX, YY)
  
  R = ((AlphaStepVal * 50) * R) + ((AlphaStepVal * (255 - 50)) * Col.R)
  G = ((AlphaStepVal * 50) * G) + ((AlphaStepVal * (255 - 50)) * Col.G)
  B = ((AlphaStepVal * 50) * B) + ((AlphaStepVal * (255 - 50)) * Col.B)
  
  Pix(3, XX, YY) = R
  Pix(2, XX, YY) = G
  Pix(1, XX, YY) = B
  
 Loop


Next

SetDIBits ODC, OBmp, 0, H, Pix(1, 1, 1), BI, 0&

End Function

Function ArtLine(SrcColor As Long, AL_Type As T_ArtLine)

Dim Col As ExRGB
Dim XX, YY

Col = ExtractRGB(SrcColor)

'On Error Resume Next

Col = ExtractRGB(SrcColor)
If AL_Type = AL_Horizontal Then

 For YY = 1 To H - 1 Step 3
   For XX = 1 To W - 1 Step 1

  R = Pix(3, XX, YY)
  G = Pix(2, XX, YY)
  B = Pix(1, XX, YY)


  R = ((AlphaStepVal * 50) * R) + ((AlphaStepVal * (255 - 50)) * Col.R)
  G = ((AlphaStepVal * 50) * G) + ((AlphaStepVal * (255 - 50)) * Col.G)
  B = ((AlphaStepVal * 50) * B) + ((AlphaStepVal * (255 - 50)) * Col.B)

  Pix(3, XX, YY) = R
  Pix(2, XX, YY) = G
  Pix(1, XX, YY) = B

   Next
  Next
 
 Else
 
  For YY = 1 To H - 1 Step 1
   For XX = 1 To W - 1 Step 3

  R = Pix(3, XX, YY)
  G = Pix(2, XX, YY)
  B = Pix(1, XX, YY)

  
  R = ((AlphaStepVal * 50) * R) + ((AlphaStepVal * (255 - 50)) * Col.R)
  G = ((AlphaStepVal * 50) * G) + ((AlphaStepVal * (255 - 50)) * Col.G)
  B = ((AlphaStepVal * 50) * B) + ((AlphaStepVal * (255 - 50)) * Col.B)
   
  Pix(3, XX, YY) = R
  Pix(2, XX, YY) = G
  Pix(1, XX, YY) = B

   Next
  Next

 
 End If

SetDIBits ODC, OBmp, 0, H, Pix(1, 1, 1), BI, 0&

End Function

Private Function ExtractRGB(Color As Long) As ExRGB

With ExtractRGB
 
  .R = Fix(Color Mod 256)
  .G = Fix((Color / 256) Mod 256)
  .B = Fix(Color / 65536)
 
End With

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "E_Filters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'                            [Yayou-Soft]®
'                       Classe Effect_Filters
'                               2004

Enum Transf_RGB

 RBG = 1
 BRG = 2
 BGR = 3
 GBR = 4
 GRB = 5

End Enum
Enum T_Digital

 D_Normal = 2
 D_Point = 3
 D_Blur = 4
 
End Enum
Public Enum T_Emboss

 E_Normal = -128
 E_Entrer = 128
 E_Simple = 0

End Enum
Enum C_T_Emboss

 C_T_Linimosite = 1
 C_T_Simple = 2

End Enum
Public Enum T_Sharpen

 Normal = 0.5
 Plus = 0.8

End Enum
Enum T_Flip

 F_Vertical = 0
 F_Horizontal = -1
 
End Enum
Function BlackAndWhite()

Dim XX, YY
Dim R, G, B

On Error Resume Next

For YY = 1 To H - 1
 For XX = 1 To W - 1

  R = Pix(3, XX, YY)
  G = Pix(2, XX, YY)
  B = Pix(1, XX, YY)

If (R >= 192) Then
 R = 255: G = 255: B = 255
Else
 If (G >= 192) Then
  R = 255: G = 255: B = 255
 Else
  If (B >= 192) Then
   R = 255: G = 255: B = 255
  Else
   R = 0: G = 0: B = 0
  End If
 End If
End If

   Pix(3, XX, YY) = R
   Pix(2, XX, YY) = G
   Pix(1, XX, YY) = B

 Next
Next

SetP

End Function
Private Function ExtractRGB(Color As Long) As ExRGB
With ExtractRGB
 
  .R = Color Mod 256
  .G = (Color / 256) Mod 256
  .B = (Color / 65536)

End With
End Function
Function GreyScale(Constrast As Boolean)
On Error Resume Next
Dim XX, YY
Dim R, G, B

If Constrast = True Then
 
 For YY = 1 To H
  For XX = 1 To W
      
   R = Pix(3, XX, YY)
   G = Pix(2, XX, YY)
   B = Pix(1, XX, YY)
   
   R = 255 - ((R + G + B) / 3)
   G = R
   B = G
   
   Pix(3, XX, YY) = R
   Pix(2, XX, YY) = G
   Pix(1, XX, YY) = B

  Next
 Next
 SetP
Else
 
 For YY = 1 To H
  For XX = 1 To W
    
   R = Pix(3, XX, YY)
   G = Pix(2, XX, YY)
   B = Pix(1, XX, YY)
   
   R = ((R + G + B) / 3)
   G = R
   B = G

   Pix(3, XX, YY) = R
   Pix(2, XX, YY) = G
   Pix(1, XX, YY) = B

  Next
 Next

SetP

End If

End Function
Property Let OperationBmp(Hbmp As Long)

OBmp = Hbmp

End Property

Property Let OperationDC(hdc As Long)

ODC = hdc

End Property
Function Digital(D_Type As T_Digital)

Dim XX, YY
Dim R As Byte, G As Byte, B As Byte
 
 For YY = 5 To H - 5 Step 3
  For XX = 5 To W - 5 Step 3
  
   'copier Pix [3, 2, 1] R, G, B
   
   R = Pix(3, XX, YY)
    Pix(3, XX + D_Type, YY) = R
    Pix(3, XX, YY + D_Type) = R
    Pix(3, XX, YY - D_Type) = R
    Pix(3, XX - D_Type, YY) = R
    
    
   G = Pix(2, XX, YY)
    Pix(2, XX + D_Type, YY) = G
    Pix(2, XX, YY + D_Type) = G
    Pix(2, XX, YY - D_Type) = G
    Pix(2, XX - D_Type, YY) = G

   
    
   B = Pix(1, XX, YY)
    Pix(1, XX + D_Type, YY) = B
    Pix(1, XX, YY + D_Type) = B
    Pix(1, XX, YY - D_Type) = B
    Pix(1, XX - D_Type, YY) = B
    
     Pix(3, XX, YY) = R
     Pix(2, XX, YY) = G
     Pix(1, XX, YY) = B
   
  Next
 Next
  
'Designer
SetDIBits ODC, OBmp, 0, H, Pix(1, 1, 1), BI, 0&

End Function
Function Emboss(E_Type As T_Emboss, Linimosite As C_T_Emboss, Optional Color As Long) As Boolean

Por = 0

Dim XX, YY
Dim R, G, B
Dim Col As ExRGB

 'On Error GoTo Er
 
 Col = ExtractRGB(Color)
 
 For YY = 1 To H - 1
  For XX = 1 To W - 1
  
   'copier Pix [3, 2, 1] R, G, B
   
   R = Pix(3, XX, YY)
   G = Pix(2, XX, YY)
   B = Pix(1, XX, YY)

  R = Abs(R - Pix(3, XX + 1, YY + 1) + E_Type)
  G = Abs(G - Pix(2, XX + 1, YY + 1) + E_Type)
  B = Abs(B - Pix(1, XX + 1, YY + 1) + E_Type)

  R = (((R + G + B) / 3) + (Col.R * 0.9)) / Linimosite
  G = (((R + G + B) / 3) + (Col.G * 0.9)) / Linimosite
  B = (((R + G + B) / 3) + (Col.B * 0.9)) / Linimosite
  
If R > 255 Then R = 255: If R < 0 Then R = 0
If G > 255 Then G = 255: If G < 0 Then G = 0
If B > 255 Then B = 255: If B < 0 Then B = 0

   Pix(3, XX, YY) = R
   Pix(2, XX, YY) = G
   Pix(1, XX, YY) = B
   

  Next
  
 Next

SetP
End Function
Private Sub SetP()

SetDIBits ODC, OBmp, 0, H, Pix(1, 1, 1), BI, 0&

ReDim Pix(0)

End Sub

Function Sharpen(S_Type As T_Sharpen) As Boolean
Por = 0
Dim XX, YY
Dim R, G, B

 For YY = 2 To H - 2
  For XX = 2 To W - 2
   
   R = Pix(3, XX, YY)
   G = Pix(2, XX, YY)
   B = Pix(1, XX, YY)

 
R = R + S_Type * (R - Pix(3, XX + 1, YY + 1))
G = G + S_Type * (G - Pix(2, XX + 1, YY + 1))
B = B + S_Type * (B - Pix(1, XX + 1, YY + 1))

If R > 255 Then R = 255: If R < 0 Then R = 0
If G > 255 Then G = 255: If G < 0 Then G = 0
If B > 255 Then B = 255: If B < 0 Then B = 0

   Pix(3, XX, YY) = Abs(R)
   Pix(2, XX, YY) = Abs(G)
   Pix(1, XX, YY) = Abs(B)
   
 
  Next
  
 Next
 
SetP
End Function

Function Gamma(ValGamma As Single) As Boolean
Por = 0
Dim XX, YY
Dim R, G, B



'On Error Resume Next
 
 For YY = 1 To H
  For XX = 1 To W
   
   'copier Pix [3, 2, 1] R, G, B
   
   R = Pix(3, XX, YY)
   G = Pix(2, XX, YY)
   B = Pix(1, XX, YY)

R = ((R / 255) ^ (1 / ValGamma)) * 255
G = ((G / 255) ^ (1 / ValGamma)) * 255
B = ((B / 255) ^ (1 / ValGamma)) * 255

   Pix(3, XX, YY) = R
   Pix(2, XX, YY) = G
   Pix(1, XX, YY) = B
   

   
  Next
 
   
 Next
SetP

End Function
Function Flame()

Dim XX, YY
Dim R, G, B

'On Error Resume Next
 
 For YY = 1 To H
  For XX = 1 To W
    
   R = Pix(3, XX, YY)
   G = Pix(2, XX, YY)
   B = Pix(1, XX, YY)

 If B >= R Then
  R = (R + ((R + G + B) / 3))
  B = Abs(R - ((R + G + B) / 3))
 End If

If R < 0 Then R = 0
If R > 255 Then R = 255
If G < 0 Then G = 0
If G > 255 Then G = 255
If B < 0 Then B = 0
If B > 255 Then B = 255

   Pix(3, XX, YY) = R
   Pix(2, XX, YY) = G
   Pix(1, XX, YY) = B
   
  Next
 Next

SetP

End Function
Function Brightness(ValBrightness)
Por = 0
Dim XX, YY
Dim R, G, B

On Error Resume Next
 
 For YY = 1 To H - 1
  For XX = 1 To W - 1
   
   Pix(3, XX, YY) = Pix(3, XX, YY) + ValBrightness
   Pix(2, XX, YY) = Pix(2, XX, YY) + ValBrightness
   Pix(1, XX, YY) = Pix(1, XX, YY) + ValBrightness
   
   If Pix(3, XX, YY) < 0 Then Pix(3, XX, YY) = 0: If Pix(3, XX, YY) > 255 Then Pix(3, XX, YY) = 255
   If Pix(2, XX, YY) < 0 Then Pix(2, XX, YY) = 0: If Pix(2, XX, YY) > 255 Then Pix(2, XX, YY) = 255
   If Pix(1, XX, YY) < 0 Then Pix(1, XX, YY) = 0: If Pix(1, XX, YY) > 255 Then Pix(1, XX, YY) = 255
   
  Next
 
   
 Next


SetP

End Function

Function Contrast(ValContrast)

Dim XX, YY
Dim R, G, B

'On Error Resume Next
 
 For YY = 1 To H - 1
  For XX = 1 To W - 1

         R = Pix(3, XX, YY)
         G = Pix(2, XX, YY)
         B = Pix(1, XX, YY)
         
 R = ((R - 128) * ValContrast) + 128
 G = ((G - 128) * ValContrast) + 128
 B = ((B - 128) * ValContrast) + 128
         
          If R < 0 Then R = 0
          If R > 255 Then R = 255
          If G < 0 Then G = 0
          If G > 255 Then G = 255
          If B < 0 Then B = 0
          If B > 255 Then B = 255
          
   Pix(3, XX, YY) = R
   Pix(2, XX, YY) = G
   Pix(1, XX, YY) = B
   
  Next
 Next

SetP

End Function

Function TransformationRGB(RGBTo As Transf_RGB)

Dim XX, YY
Dim R, G, B

 For YY = 1 To H - 1
  For XX = 1 To W - 1

   R = Pix(3, XX, YY)
   G = Pix(2, XX, YY)
   B = Pix(1, XX, YY)
   
   Select Case RGBTo
     
     Case 1: R = Pix(3, XX, YY): B = Pix(1, XX, YY): G = Pix(2, XX, YY)
     
     Case 2: R = Pix(1, XX, YY): B = Pix(3, XX, YY): G = Pix(2, XX, YY)
     
     Case 3: R = Pix(1, XX, YY): B = Pix(2, XX, YY): G = Pix(3, XX, YY)
    
     Case 4: R = Pix(2, XX, YY): B = Pix(1, XX, YY): G = Pix(3, XX, YY)
    
     Case 5: R = Pix(2, XX, YY): B = Pix(3, XX, YY): G = Pix(1, XX, YY)
    
    End Select
   
   Pix(3, XX, YY) = R
   Pix(2, XX, YY) = G
   Pix(1, XX, YY) = B
    
  Next
 Next

SetP

End Function

Private Sub Class_Terminate()

ReDim Pix(0)

End Sub
Function Cilibert(AddR, AddG, AddB)

Dim XX, YY
Dim R, G, B

 For YY = 1 To H - 1
  For XX = 1 To W - 1

   R = Pix(3, XX, YY) + AddR
   G = Pix(2, XX, YY) + AddG
   B = Pix(1, XX, YY) + AddB
         
  If R < 0 Then R = 0
  If R > 255 Then R = 255
  
  If G < 0 Then G = 0
  If G > 255 Then G = 255
  
  If B < 0 Then B = 0
  If B > 254 Then B = 255
          
   Pix(3, XX, YY) = R
   Pix(2, XX, YY) = G
   Pix(1, XX, YY) = B




  Next
 Next
 
 SetP
 
End Function
Sub Loadning(hdc As Long, Hbmp As Long, WD, HG)

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
 
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dimentions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Function Flip(T_Flip As T_Flip)

If T_Flip = F_Horizontal Then

  StretchBlt ODC, W, 0, -W, H, ODC, 0, 0, W, H, vbSrcCopy
  
Else

StretchBlt ODC, 0, H, W, -H, ODC, 0, 0, W, H, vbSrcCopy

End If


End Function

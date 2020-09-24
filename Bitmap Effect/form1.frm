VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Bitmap Effect 1.1  [Yayou-Soft®]"
   ClientHeight    =   4965
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7215
   Icon            =   "form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3240
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "fichier bitmap (*.Bmp)|*.Bmp"
   End
   Begin VB.PictureBox PD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   4080
      ScaleHeight     =   214
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   184
      TabIndex        =   1
      Top             =   120
      Width           =   2760
   End
   Begin VB.PictureBox PS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   120
      Picture         =   "form1.frx":0442
      ScaleHeight     =   214
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   184
      TabIndex        =   0
      Top             =   120
      Width           =   2760
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Yayou-Soft®"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   3600
      Width           =   6270
   End
   Begin VB.Menu MnuFilters 
      Caption         =   "Filters"
      Begin VB.Menu MnuDigital 
         Caption         =   "Digital"
      End
      Begin VB.Menu MnuEm 
         Caption         =   "Emboss"
         Begin VB.Menu MnuEmboss 
            Caption         =   "Emboss Simple"
         End
         Begin VB.Menu MnuEmE 
            Caption         =   "Emboss entrer"
         End
         Begin VB.Menu MnuEN 
            Caption         =   "Emboss normal"
         End
      End
      Begin VB.Menu MnuSharpen 
         Caption         =   "Sharpen"
      End
      Begin VB.Menu MnuGamma 
         Caption         =   "Gamma"
      End
      Begin VB.Menu MnuBrightness 
         Caption         =   "Brightness"
      End
      Begin VB.Menu MnuContrast 
         Caption         =   "Contrast"
      End
      Begin VB.Menu MnuEclilibre 
         Caption         =   "Eclibrer les couleur"
      End
      Begin VB.Menu MnuGris 
         Caption         =   "Niveaux de gris"
      End
      Begin VB.Menu MnuNoire 
         Caption         =   "Noire et blanc"
      End
      Begin VB.Menu MnuFlame 
         Caption         =   "Flame"
      End
      Begin VB.Menu MnuTransformation 
         Caption         =   "Transformation RGB"
      End
      Begin VB.Menu MnuBlur 
         Caption         =   "Blur"
      End
      Begin VB.Menu MnuMosaic 
         Caption         =   "Mosaic"
      End
   End
   Begin VB.Menu MnuEdition 
      Caption         =   "Edition"
      Begin VB.Menu MnuCopier 
         Caption         =   "Copier"
      End
      Begin VB.Menu MnuEffacer 
         Caption         =   "Effacer"
      End
      Begin VB.Menu MnuSp 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOuvrir 
         Caption         =   "Ouvrir"
      End
      Begin VB.Menu MnuEnregestrer 
         Caption         =   "Enregestrer"
      End
      Begin VB.Menu MnuSp2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuImport 
         Caption         =   "Import"
      End
      Begin VB.Menu MnuExporter 
         Caption         =   "Exporter [format bmp]"
      End
   End
   Begin VB.Menu MnuArt 
      Caption         =   "Art"
      Begin VB.Menu MnuArtLine 
         Caption         =   "Art Ligne"
         Begin VB.Menu MnuVertical 
            Caption         =   "Vertical"
         End
         Begin VB.Menu MnuHorisontal 
            Caption         =   "Horizontal"
         End
      End
      Begin VB.Menu MnuArtCercle 
         Caption         =   "Art Cercle"
      End
   End
   Begin VB.Menu MnuDraw 
      Caption         =   "Draw"
      Begin VB.Menu MnuCircle 
         Caption         =   "Circle"
      End
      Begin VB.Menu MnuText 
         Caption         =   "Text"
      End
      Begin VB.Menu MnuCouver 
         Caption         =   "Couver"
      End
   End
   Begin VB.Menu MnuDimontion 
      Caption         =   "Dimontion"
      Begin VB.Menu MnuFlip 
         Caption         =   "Flip"
         Begin VB.Menu MnuFHorizontal 
            Caption         =   "Horizontal"
         End
         Begin VB.Menu MnuFVertical 
            Caption         =   "Vertical"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EF As New E_Filters
Dim ED As New Edtion
Dim Ar As New E_Arts
Dim Dr As New Draws
Dim DM As New Dimentions

Private Sub Form_Activate()

 Show
 EF.OperationDC = PD.hdc
 EF.OperationBmp = PD.Image

 ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight

End Sub

Private Sub MnuArtCercle_Click()

EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
'-------------------------------------------------------------

Ar.ArtCircle PD.ScaleWidth / 2, PD.ScaleHeight / 2, vbRed
PD.Picture = PD.Image

End Sub

Private Sub MnuBlur_Click()

EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
EF.MonsionBlur

End Sub

Private Sub MnuBrightness_Click()

 EF.OperationDC = PD.hdc
 EF.OperationBmp = PD.Image

 ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
 '-------------------------------------------------------------
Msg = InputBox("Entrer La valeur de -255 jusca 255")

 EF.Brightness Val(Msg)
 PD.Picture = PD.Image

End Sub
Private Sub MnuCircle_Click()

 Dr.DrawCircle PD.ScaleWidth / 2, PD.ScaleHeight / 2, 50, vbGreen
 PD.Picture = PD.Image

End Sub
Private Sub MnuContrast_Click()

 EF.OperationDC = PD.hdc
 EF.OperationBmp = PD.Image

 ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
 '-------------------------------------------------------------

 Valeur = InputBox("Entrer la valeur :" & vbCrLf & "Min 0 - max 5)")

 EF.Contrast Valeur
 PD.Picture = PD.Image

End Sub
Private Sub MnuCopier_Click()

 ED.CopyImage
 
 End Sub
Private Sub MnuCouver_Click()

 'Dr.DrawCouver
 PD.Picture = PD.Image

End Sub
Private Sub MnuDigital_Click()

 EF.OperationDC = PD.hdc
 EF.OperationBmp = PD.Image
 
 ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
 
 '-------------------------------------------------------------
 Msg = InputBox("choiser 1, 2, 3, ...:  " & vbCrLf & "1-Blur" & vbCrLf & "2-Normal" & vbCrLf & "3-Point")
 
 Select Case Msg
 
  Case 1: EF.Digital D_Blur
  Case 2: EF.Digital D_Normal
  Case 3: EF.Digital D_Point
 
 End Select
 
 PD.Picture = PD.Image
 
 
End Sub
Private Sub MnuEclilibre_Click()

Dim R, G, B

EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

 ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
'-------------------------------------------------------------
R = InputBox("Entrer valeur de Rouge")
G = InputBox("Entrer valeur de Vert")
B = InputBox("Entrer valeur de Blur")

 EF.Cilibert R, G, B
 
PD.Picture = PD.Image

End Sub

Private Sub MnuEffacer_Click()
ED.Clear

PD.Picture = LoadPicture()

End Sub

Private Sub MnuEmboss_Click()
EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
'-------------------------------------------------------------
Reponse = MsgBox("Linimosite", vbYesNo)
If Reponse = vbYes Then
 EF.Emboss E_Normal, C_T_Linimosite, vbBlue
Else
 EF.Emboss E_Normal, C_T_Simple, vbBlue
 
End If

PD.Picture = PD.Image

End Sub


Private Sub MnuEmE_Click()

EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
'-------------------------------------------------------------
Reponse = MsgBox("Linimosite", vbYesNo)
If Reponse = vbYes Then
 EF.Emboss E_Entrer, C_T_Linimosite, vbRed
Else
 EF.Emboss E_Entrer, C_T_Simple, vbRed
 
End If

PD.Picture = PD.Image


End Sub


Private Sub MnuEN_Click()

EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
'-------------------------------------------------------------
Reponse = MsgBox("Linimosite", vbYesNo)
If Reponse = vbYes Then
 EF.Emboss E_Simple, C_T_Linimosite, vbRed
Else
 EF.Emboss E_Simple, C_T_Simple, vbRed
 
End If

PD.Picture = PD.Image

End Sub


Private Sub MnuEnregestrer_Click()
CD1.Flags = cdlOFNHideReadOnly

CD1.Filter = "fichier Yayou-Soft® Picture (*.Ysp)|*.Ysp|"

CD1.ShowSave

ED.addObject CD1.FileName

End Sub

Private Sub MnuExporter_Click()

CD1.Flags = cdlOFNHideReadOnly

CD1.Filter = "fichier Yayou-Soft® Picture (*.bmp)|*.bmp|"

CD1.ShowSave

SavePicture PD.Picture, CD1.FileName & ".Bmp"



End Sub

Private Sub MnuFHorizontal_Click()

EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
'-------------------------------------------------------------

DM.Flip F_Horizontal

End Sub

Private Sub MnuFlame_Click()

EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
'-------------------------------------------------------------

EF.Flame

PD.Picture = PD.Image

End Sub

Private Sub MnuFVertical_Click()

EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
'-------------------------------------------------------------

DM.Flip F_Vertical


End Sub


Private Sub MnuGamma_Click()

EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
'-------------------------------------------------------------
Msg = InputBox("Entrer la valeur de Gamma  de 0 to 5")
EF.Gamma Val(Msg)

PD.Picture = PD.Image

End Sub
Private Sub MnuGris_Click()

EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

 ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
'-------------------------------------------------------------
Msg = MsgBox("Linimosite", vbYesNo)

If Msg = vbYes Then

 EF.GreyScale True
 
Else

EF.GreyScale False

End If

PD.Picture = PD.Image

End Sub

Private Sub MnuHorisontal_Click()
EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
'-------------------------------------------------------------
Ar.ArtLine RGB(125, 125, 125), AL_Horizontal

End Sub

Private Sub MnuImport_Click()
CD1.Flags = cdlOFNHideReadOnly

CD1.Filter = "Tous les format de image |*.bmp;*.jpg;*.gif;*.png|Tous les fichier(*.*)|*.*|"

CD1.ShowOpen


PS.Picture = LoadPicture(CD1.FileName)

End Sub

Private Sub MnuMosaic_Click()

EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight

EF.Mosaic

End Sub

Private Sub MnuNoire_Click()


EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
'-------------------------------------------------------------

EF.BlackAndWhite

PD.Picture = PD.Image

End Sub

Private Sub MnuOuvrir_Click()
CD1.Flags = cdlOFNHideReadOnly

CD1.Filter = "fichier Yayou-Soft® Picture (*.Ysp)|*.Ysp|Tous les fichier(*.*)|*.*|"

CD1.ShowOpen


ED.LoadObject CD1.FileName, 2

PD.Picture = PD.Image
'MsgBox "ok"

End Sub


Private Sub MnuSharpen_Click()

EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

 ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
 '-------------------------------------------------------------

 EF.Sharpen Plus
 PD.Picture = PD.Image

End Sub

Private Sub MnuText_Click()
'PD.Cls
Dim Msg As String

Msg = InputBox("Entrer le text ici :")

Dr.WrawText Msg, PD.ScaleWidth / 2, PD.ScaleHeight / 2, vbRed, True
PD.Picture = PD.Image


End Sub

Private Sub MnuTransformation_Click()
EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
'-------------------------------------------------------------
Msg = InputBox("Choiser 1, 2, 3, ..." & vbCrLf & "1-RBG" & vbCrLf & "2-BRG" & vbCrLf & "3-BGR" & vbCrLf & "4-GBR" & vbCrLf & "5-GRB")

EF.TransformationRGB Val(Msg)
PD.Picture = PD.Image

End Sub

Private Sub MnuVertical_Click()
EF.OperationDC = PD.hdc
EF.OperationBmp = PD.Image

ED.Loading PS.hdc, PS.Image, PS.ScaleWidth, PS.ScaleHeight
'-------------------------------------------------------------

Ar.ArtLine &HFFFFFF, AL_Vertical

End Sub

Private Sub Timer1_Timer()

Caption = "fps   :  " & FPS
FPS = 0
End Sub

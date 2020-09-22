VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFilters 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Filters"
   ClientHeight    =   -30
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4500
   Icon            =   "Filters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   -2
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Unten ausrichten
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   -255
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   397
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   15
      Left            =   0
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   361
      TabIndex        =   1
      Top             =   0
      Width           =   5415
   End
   Begin VB.PictureBox picImagem 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   15
      Left            =   0
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6045
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All types |*.*"
      FontSize        =   2,54052e-29
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu mnuOpenFile 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuSaveFile 
         Caption         =   "Save..."
      End
      Begin VB.Menu mnuSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuFilters 
      Caption         =   "Filters"
      Begin VB.Menu blur 
         Caption         =   "Blur"
         Begin VB.Menu mnuBlur 
            Caption         =   "Blur..."
         End
         Begin VB.Menu mnuBlurMore 
            Caption         =   "Blur More..."
         End
      End
      Begin VB.Menu Sharpen 
         Caption         =   "Sharpen"
         Begin VB.Menu mnuSharpen 
            Caption         =   "Sharpen..."
         End
         Begin VB.Menu mnuSharpenMore 
            Caption         =   "Sharpen More..."
         End
      End
      Begin VB.Menu mnuEmboss 
         Caption         =   "Emboss..."
      End
      Begin VB.Menu Difuse 
         Caption         =   "Difuse"
         Begin VB.Menu mnuDifuse 
            Caption         =   "Difuse..."
         End
         Begin VB.Menu mnuDifuseMore 
            Caption         =   "Difuse More..."
         End
      End
      Begin VB.Menu mnuDummy 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBubbles 
         Caption         =   "Circles"
      End
      Begin VB.Menu mnuBubbles2 
         Caption         =   "Bubbles"
      End
   End
End
Attribute VB_Name = "frmFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, j
Dim Imagem
Dim AletatorioX, AletatorioY
Dim Pixel&
Dim Red As Integer, Green As Integer, Blue As Integer




Private Sub Form_Resize()
    pBar.Top = ScaleHeight - pBar.Height - 2
End Sub

Private Sub mnuBubbles_Click()
    On Error Resume Next
    Static TmpRAD As Single
    Dim xx As Integer, yy As Integer
    
    TmpRAD = 40
    
    For yy = 0 To Picture1.ScaleHeight
        For xx = 0 To Picture1.Width
            Picture1.Circle (xx, yy), TmpRAD, Picture1.Point(xx, yy)
            xx = xx + 10
        Next xx
        yy = yy + 10
        frmFilters.pBar.Value = (((Picture1.ScaleHeight - 1) / 100) * yy)
        sBar.SimpleText = "Draw(" & Trim(Str(Int(pBar.Value))) & "%)"
    Next yy
    Picture1.Refresh
    frmFilters.pBar.Value = 0
    
    sBar.SimpleText = ""
End Sub

Private Sub mnuBubbles2_Click()
Static xx, yy As Single
Picture1.DrawWidth = 15
For yy = 0 To Picture1.ScaleHeight
For xx = 0 To Picture1.ScaleWidth
Picture1.PSet (xx, yy), Picture1.Point(xx, yy)
xx = xx + 15
Next xx
yy = yy + 15
Next yy
Picture1.DrawWidth = 1

End Sub


Private Sub mnuOpenFile_Click()
CommonDialog1.Action = 1
Imagem = CommonDialog1.FileName
If Imagem = "" Then Exit Sub
Picture1.Picture = LoadPicture(Imagem)
frmFilters.Refresh

x = Picture1.ScaleWidth
y = Picture1.ScaleHeight
If x > 750 Or y > 750 Then
    MsgBox "Please, choose a smaller picture"
    x = 0
    y = 0
    Exit Sub
End If
'resizing the form....
frmFilters.Width = frmFilters.ScaleX(Picture1.Width + 6, vbPixels, vbTwips)
frmFilters.Height = frmFilters.ScaleY(Picture1.Height + 30, vbPixels, vbTwips)
frmFilters.Refresh
End Sub

Private Sub mnuSaveFile_Click()
    CommonDialog1.Action = 2
    Imagem = CommonDialog1.FileName
    SavePicture Picture1.Image, Imagem
End Sub

Private Sub mnuSharpenMore_Click()

Call loading(i, j)
    For i = 1 To y - 2
        For j = 1 To x - 2
            Red = ImageArray(0, i, j) + 0.8 * (ImageArray(0, i, j) - ImageArray(0, i - 1, j - 1))
            Green = ImageArray(1, i, j) + 0.8 * (ImageArray(1, i, j) - ImageArray(1, i - 1, j - 1))
            Blue = ImageArray(2, i, j) + 0.8 * (ImageArray(2, i, j) - ImageArray(2, i - 1, j - 1))
            If Red > 255 Then Red = 255
            If Red < 0 Then Red = 0
            If Green > 255 Then Green = 255
            If Green < 0 Then Green = 0
            If Blue > 255 Then Blue = 255
            If Blue < 0 Then Blue = 0
            SetPixelV Picture1.hdc, j, i, RGB(Red, Green, Blue)
        Next
        frmFilters.pBar.Value = i * 100 / (y - 1)
                sBar.SimpleText = "Draw(" & Trim(Str(Int(pBar.Value))) & "%)"
        
    Next
    pBar.Value = 0
    frmFilters.sBar.SimpleText = ""
End Sub

Private Sub mnuSharpen_Click()
   
   Call loading(i, j)
    For i = 1 To y - 2
        For j = 1 To x - 2
            Red = ImageArray(0, i, j) + 0.5 * (ImageArray(0, i, j) - ImageArray(0, i - 1, j - 1))
            Green = ImageArray(1, i, j) + 0.5 * (ImageArray(1, i, j) - ImageArray(1, i - 1, j - 1))
            Blue = ImageArray(2, i, j) + 0.5 * (ImageArray(2, i, j) - ImageArray(2, i - 1, j - 1))
            If Red > 255 Then Red = 255
            If Red < 0 Then Red = 0
            If Green > 255 Then Green = 255
            If Green < 0 Then Green = 0
            If Blue > 255 Then Blue = 255
            If Blue < 0 Then Blue = 0
            SetPixelV Picture1.hdc, j, i, RGB(Red, Green, Blue)
        Next
        frmFilters.pBar.Value = i * 100 / (y - 1)
                sBar.SimpleText = "Draw(" & Trim(Str(Int(pBar.Value))) & "%)"
    Next
    pBar.Value = 0
    frmFilters.sBar.SimpleText = ""
End Sub

Private Sub mnuEmboss_Click()
   
   Call loading(i, j)

    For i = 1 To y - 2
        For j = 1 To x - 2
            Red = Abs(ImageArray(0, i, j) - ImageArray(0, i + 1, j + 1) + 128)
            Green = Abs(ImageArray(1, i, j) - ImageArray(1, i + 1, j + 1) + 128)
            Blue = Abs(ImageArray(2, i, j) - ImageArray(2, i + 1, j + 1) + 128)
            SetPixelV Picture1.hdc, j, i, RGB(Red, Green, Blue)
        Next
        frmFilters.pBar.Value = i * 100 / (y - 1)
                sBar.SimpleText = "Draw(" & Trim(Str(Int(pBar.Value))) & "%)"
    Next
    pBar.Value = 0
    frmFilters.sBar.SimpleText = ""
End Sub

Private Sub mnuDifuseMore_Click()

Call loading(i, j)
For i = 2 To y - 2
    For j = 2 To x - 2
        AletatorioX = Rnd * 6 - 2
        AletatorioY = Rnd * 6 - 2
        Red = ImageArray(0, i + AletatorioX, j + AletatorioY)
        Green = ImageArray(1, i + AletatorioX, j + AletatorioY)
        Blue = ImageArray(2, i + AletatorioX, j + AletatorioY)
        SetPixelV Picture1.hdc, j, i, RGB(Red, Green, Blue)
    Next
        frmFilters.pBar.Value = i * 100 / (y - 1)
                sBar.SimpleText = "Draw(" & Trim(Str(Int(pBar.Value))) & "%)"
    Next
    pBar.Value = 0
    frmFilters.sBar.SimpleText = ""
End Sub

Private Sub mnuBlur_Click()
    
Call loading(i, j)
    For i = 1 To y - 2
        For j = 1 To x - 2
            Red = ImageArray(0, i - 1, j - 1) + ImageArray(0, i - 1, j) + ImageArray(0, i - 1, j + 1) + _
            ImageArray(0, i, j - 1) + ImageArray(0, i, j) + ImageArray(0, i, j + 1) + _
            ImageArray(0, i + 1, j - 1) + ImageArray(0, i + 1, j) + ImageArray(0, i + 1, j + 1)
            
            
            
            Green = ImageArray(1, i - 1, j - 1) + ImageArray(1, i - 1, j) + ImageArray(1, i - 1, j + 1) + _
            ImageArray(1, i, j - 1) + ImageArray(1, i, j) + ImageArray(1, i, j + 1) + _
            ImageArray(1, i + 1, j - 1) + ImageArray(1, i + 1, j) + ImageArray(1, i + 1, j + 1)

            
            Blue = ImageArray(2, i - 1, j - 1) + ImageArray(2, i - 1, j) + ImageArray(2, i - 1, j + 1) + _
            ImageArray(2, i, j - 1) + ImageArray(2, i, j) + ImageArray(2, i, j + 1) + _
            ImageArray(2, i + 1, j - 1) + ImageArray(2, i + 1, j) + ImageArray(2, i + 1, j + 1)

            SetPixelV Picture1.hdc, j, i, RGB(Red / 9, Green / 9, Blue / 9)
        Next
        frmFilters.pBar.Value = i * 100 / (y - 1)
                sBar.SimpleText = "Draw(" & Trim(Str(Int(pBar.Value))) & "%)"
    Next
    pBar.Value = 0
    frmFilters.sBar.SimpleText = ""
End Sub

Private Sub mnuDifuse_Click()

Call loading(i, j)

    For i = 2 To y - 2
        For j = 2 To x - 2
            AletatorioX = Rnd * 3 - 2
            AletatorioY = Rnd * 3 - 2
            Red = ImageArray(0, i + AletatorioX, j + AletatorioY)
            Green = ImageArray(1, i + AletatorioX, j + AletatorioY)
            Blue = ImageArray(2, i + AletatorioX, j + AletatorioY)
            SetPixelV Picture1.hdc, j, i, RGB(Red, Green, Blue)
        Next
        frmFilters.pBar.Value = i * 100 / (y - 1)
        sBar.SimpleText = "Draw(" & Trim(Str(Int(pBar.Value))) & "%)"
    Next
    pBar.Value = 0
    frmFilters.sBar.SimpleText = ""
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuBlurMore_Click()

Call loading(i, j)

    For i = 1 To y - 2
        For j = 1 To x - 2
            Red = ImageArray(0, i - 2, j - 2) + ImageArray(0, i - 2, j - 1) + ImageArray(0, i - 2, j) + _
            ImageArray(0, i - 2, j + 1) + ImageArray(0, i - 2, j + 2) + ImageArray(0, i - 1, j - 2) + _
            ImageArray(0, i - 1, j - 1) + ImageArray(0, i - 1, j) + ImageArray(0, i - 1, j + 1) + _
            ImageArray(0, i - 1, j + 2) + ImageArray(0, i, j - 2) + _
            ImageArray(0, i, j - 1) + ImageArray(0, i, j) + ImageArray(0, i, j + 1) + _
            ImageArray(0, i, j + 2) + ImageArray(0, i + 1, j - 2) + _
            ImageArray(0, i + 1, j - 1) + ImageArray(0, i + 1, j) + ImageArray(0, i + 1, j + 1) + _
            ImageArray(0, i + 1, j + 2) + ImageArray(0, i + 2, j - 2) + ImageArray(0, i + 2, j - 1) + _
            ImageArray(0, i + 2, j) + ImageArray(0, i + 2, j + 1) + ImageArray(0, i + 2, j + 2)
            
            
            Green = ImageArray(1, i - 2, j - 2) + ImageArray(1, i - 2, j - 1) + ImageArray(1, i - 2, j) + _
            ImageArray(1, i - 2, j + 1) + ImageArray(1, i - 2, j + 2) + ImageArray(1, i - 1, j - 2) + _
            ImageArray(1, i - 1, j - 1) + ImageArray(1, i - 1, j) + ImageArray(1, i - 1, j + 1) + _
            ImageArray(1, i - 1, j + 2) + ImageArray(1, i, j - 2) + _
            ImageArray(1, i, j - 1) + ImageArray(1, i, j) + ImageArray(1, i, j + 1) + _
            ImageArray(1, i, j + 2) + ImageArray(1, i + 1, j - 2) + _
            ImageArray(1, i + 1, j - 1) + ImageArray(1, i + 1, j) + ImageArray(1, i + 1, j + 1) + _
            ImageArray(1, i + 1, j + 2) + ImageArray(1, i + 2, j - 2) + ImageArray(1, i + 2, j - 1) + _
            ImageArray(1, i + 2, j) + ImageArray(1, i + 2, j + 1) + ImageArray(1, i + 2, j + 2)
            
            Blue = ImageArray(2, i - 2, j - 2) + ImageArray(2, i - 2, j - 1) + ImageArray(2, i - 2, j) + _
            ImageArray(2, i - 2, j + 1) + ImageArray(2, i - 2, j + 2) + ImageArray(2, i - 1, j - 2) + _
            ImageArray(2, i - 1, j - 1) + ImageArray(2, i - 1, j) + ImageArray(2, i - 1, j + 1) + _
            ImageArray(2, i - 1, j + 2) + ImageArray(2, i, j - 2) + _
            ImageArray(2, i, j - 1) + ImageArray(2, i, j) + ImageArray(2, i, j + 1) + _
            ImageArray(2, i, j + 2) + ImageArray(2, i + 1, j - 2) + _
            ImageArray(2, i + 1, j - 1) + ImageArray(2, i + 1, j) + ImageArray(2, i + 1, j + 1) + _
            ImageArray(2, i + 1, j + 2) + ImageArray(2, i + 2, j - 2) + ImageArray(2, i + 2, j - 1) + _
            ImageArray(2, i + 2, j) + ImageArray(2, i + 2, j + 1) + ImageArray(2, i + 2, j + 2)

            SetPixelV Picture1.hdc, j, i, RGB(Red / 25, Green / 25, Blue / 25)
        Next
        frmFilters.pBar.Value = i * 100 / (y - 1)
        sBar.SimpleText = "Draw(" & Trim(Str(Int(pBar.Value))) & "%)"
    Next
    pBar.Value = 0
    frmFilters.sBar.SimpleText = ""
End Sub



Attribute VB_Name = "ImgMod"
Global ImageArray(-4 To 2, -4 To 700, -4 To 700) As Integer
Global x, y As Integer

Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long


Sub loading(i, j)

    Dim Color As Long
    frmFilters.sBar.SimpleText = "Calc"
    
    For i = 0 To y - 1
        For j = 0 To x - 1
            Pixel& = frmFilters.Picture1.Point(j, i)
            Red = Pixel& Mod 256
            Green = ((Pixel& And &HFF00) / 256&) Mod 256&
            Blue = (Pixel& And &HFF0000) / 65536
            ImageArray(0, i, j) = Red
            ImageArray(1, i, j) = Green
            ImageArray(2, i, j) = Blue
        Next
        
        frmFilters.pBar.Value = i * 100 / (y - 1)
    Next
    frmFilters.pBar.Value = 0
    frmFilters.sBar.SimpleText = "Draw"
End Sub




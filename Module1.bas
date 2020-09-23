Attribute VB_Name = "Module1"
Option Explicit
'Code by: Dan Redding - Blue Knot Software
'For the FULL version of this module, please visit
' http://www.planet-source-code.com/vb
'(The darken & brighten routines in this module are
'slightly modified from that version)

'Portions of this code marked with *** are converted from
'C/C++ routines for RGB/HSL conversion found in the
'Microsoft Knowledge Base (PD sample code):
'http://support.microsoft.com/support/kb/articles/Q29/2/40.asp
'In addition to the language conversion, some internal
'calculations have been modified and converted to FP math to
'reduce rounding errors.
'Conversion to VB and original code by
'Dan Redding (bwsoft@revealed.net)
'http://home.revealed.net/bwsoft
'Free to use, please give proper credit
Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Const HSLMAX As Integer = 240 '***
    'H, S and L values can be 0 - HSLMAX
    '240 matches what is used by MS Win;
    'any number less than 1 byte is OK;
    'works best if it is evenly divisible by 6
Const RGBMAX As Integer = 255 '***
    'R, G, and B value can be 0 - RGBMAX
Const UNDEFINED As Integer = (HSLMAX * 2 / 3) '***
    'Hue is undefined if Saturation = 0 (greyscale)

Public Type HSLCol 'Datatype used to pass HSL Color values
    Hue As Integer
    Sat As Integer
    Lum As Integer
End Type

Public Function RGBRed(RGBCol As Long) As Integer
'Return the Red component from an RGB Color
    RGBRed = RGBCol And &HFF
End Function

Public Function RGBGreen(RGBCol As Long) As Integer
'Return the Green component from an RGB Color
    RGBGreen = ((RGBCol And &H100FF00) / &H100)
End Function

Public Function RGBBlue(RGBCol As Long) As Integer
'Return the Blue component from an RGB Color
    RGBBlue = (RGBCol And &HFF0000) / &H10000
End Function

Public Function WebSafe(intVal As Variant) As Integer
' I did not write this part found on the net somewere can't remmber'
' just like to say thanks for whoever did.
    Select Case intVal
        Case 0, 51, 102, 153, 204, 255
            WebSafe = intVal
        Case Else
            If intVal <= 26 Then
                WebSafe = 0: Exit Function
            ElseIf intVal > 26 And intVal <= 76 Then
                WebSafe = 51: Exit Function
            ElseIf intVal > 76 And intVal <= 127 Then
                WebSafe = 102: Exit Function
            ElseIf intVal > 127 And intVal <= 178 Then
                WebSafe = 153: Exit Function
            ElseIf intVal > 178 And intVal <= 229 Then
                WebSafe = 204: Exit Function
            ElseIf intVal > 229 Then
                WebSafe = 255: Exit Function
            End If
    End Select
End Function

Private Function iMax(a As Integer, b As Integer) _
    As Integer
'Return the Larger of two values
    iMax = IIf(a > b, a, b)
End Function

Private Function iMin(a As Integer, b As Integer) _
    As Integer
'Return the smaller of two values
    iMin = IIf(a < b, a, b)
End Function

Public Function RGBtoHSL(RGBCol As Long) As HSLCol '***
'Returns an HSLCol datatype containing Hue, Luminescence
'and Saturation; given an RGB Color value

Dim r As Integer, g As Integer, b As Integer
Dim cMax As Integer, cMin As Integer
Dim RDelta As Double, GDelta As Double, _
    BDelta As Double
Dim H As Double, s As Double, L As Double
Dim cMinus As Long, cPlus As Long
    
    r = RGBRed(RGBCol)
    g = RGBGreen(RGBCol)
    b = RGBBlue(RGBCol)
    
    cMax = iMax(iMax(r, g), b) 'Highest and lowest
    cMin = iMin(iMin(r, g), b) 'color values
    
    cMinus = cMax - cMin 'Used to simplify the
    cPlus = cMax + cMin  'calculations somewhat.
    
    'Calculate luminescence (lightness)
    L = ((cPlus * HSLMAX) + RGBMAX) / (2 * RGBMAX)
    
    If cMax = cMin Then 'achromatic (r=g=b, greyscale)
        s = 0 'Saturation 0 for greyscale
        H = UNDEFINED 'Hue undefined for greyscale
    Else
        'Calculate color saturation
        If L <= (HSLMAX / 2) Then
            s = ((cMinus * HSLMAX) + 0.5) / cPlus
        Else
            s = ((cMinus * HSLMAX) + 0.5) / (2 * RGBMAX - cPlus)
        End If
    
        'Calculate hue
        RDelta = (((cMax - r) * (HSLMAX / 6)) + 0.5) / cMinus
        GDelta = (((cMax - g) * (HSLMAX / 6)) + 0.5) / cMinus
        BDelta = (((cMax - b) * (HSLMAX / 6)) + 0.5) / cMinus
    
        Select Case cMax
            Case CLng(r)
                H = BDelta - GDelta
            Case CLng(g)
                H = (HSLMAX / 3) + RDelta - BDelta
            Case CLng(b)
                H = ((2 * HSLMAX) / 3) + GDelta - RDelta
        End Select
        
        If H < 0 Then H = H + HSLMAX
    End If
    
    RGBtoHSL.Hue = CInt(H)
    RGBtoHSL.Lum = CInt(L)
    RGBtoHSL.Sat = CInt(s)
End Function

Public Function HSLtoRGB(HueLumSat As HSLCol) As Long '***
    Dim r As Double, g As Double, b As Double
    Dim H As Double, L As Double, s As Double
    Dim Magic1 As Double, Magic2 As Double
    
    H = HueLumSat.Hue
    L = HueLumSat.Lum
    s = HueLumSat.Sat
    
    If CInt(s) = 0 Then 'Greyscale
        r = (L * RGBMAX) / HSLMAX 'luminescence,
                'converted to the proper range
        g = r 'All RGB values same in greyscale
        b = r
        If CInt(H) <> UNDEFINED Then
            'This is technically an error.
            'The RGBtoHSL routine will always return
            'Hue = UNDEFINED (160 when HSLMAX is 240)
            'when Sat = 0.
            'if you are writing a color mixer and
            'letting the user input color values,
            'you may want to set Hue = UNDEFINED
            'in this case.
        End If
    Else
        'Get the "Magic Numbers"
        If L <= HSLMAX / 2 Then
            Magic2 = (L * (HSLMAX + s) + 0.5) / HSLMAX
        Else
            Magic2 = L + s - ((L * s) + 0.5) / HSLMAX
        End If
        
        Magic1 = 2 * L - Magic2
        
        'get R, G, B; change units from HSLMAX range
        'to RGBMAX range
        r = (HuetoRGB(Magic1, Magic2, H + (HSLMAX / 3)) _
            * RGBMAX + 0.5) / HSLMAX
        g = (HuetoRGB(Magic1, Magic2, H) * RGBMAX + 0.5) / HSLMAX
        b = (HuetoRGB(Magic1, Magic2, H - (HSLMAX / 3)) _
            * RGBMAX + 0.5) / HSLMAX
        
    End If
    
    HSLtoRGB = RGB(CInt(r), CInt(g), CInt(b))
    
End Function

Private Function HuetoRGB(mag1 As Double, mag2 As Double, _
    ByVal Hue As Double) As Double '***
'Utility function for HSLtoRGB

'Range check
    If Hue < 0 Then
        Hue = Hue + HSLMAX
    ElseIf Hue > HSLMAX Then
        Hue = Hue - HSLMAX
    End If
    
    'Return r, g, or b value from parameters
    Select Case Hue 'Values get progressively larger.
                'Only the first true condition will execute
        Case Is < (HSLMAX / 6)
            HuetoRGB = (mag1 + (((mag2 - mag1) * Hue + _
                (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Is < (HSLMAX / 2)
            HuetoRGB = mag2
        Case Is < (HSLMAX * 2 / 3)
            HuetoRGB = (mag1 + (((mag2 - mag1) * _
                ((HSLMAX * 2 / 3) - Hue) + _
                (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Else
            HuetoRGB = mag1
    End Select
End Function

Public Function ContrastingColor(RGBCol As Long) As Long
'Returns Black or White, whichever will show up better
'on the specified color.
'Useful for setting label forecolors with transparent
'backgrounds (send it the form backcolor - RGB value, not
'system value!)
'(also produces a monochrome negative when applied to
'all pixels in an image)

Dim hsl As HSLCol
    hsl = RGBtoHSL(RGBCol)
    If hsl.Lum > HSLMAX / 2 Then ContrastingColor = 0 _
        Else: ContrastingColor = &HFFFFFF
End Function

Public Function Brighten(RGBColor As Long, Percent As Single)
'Lightens the color by a specifie percent, given as a Single
'(10% = .10)

Dim hsl As HSLCol, L As Long
    If Percent <= 0 Then
        Brighten = RGBColor
        Exit Function
    End If
    
    hsl = RGBtoHSL(RGBColor)
    L = hsl.Lum + (HSLMAX * Percent)
    If L > HSLMAX Then L = HSLMAX
    hsl.Lum = L
    Brighten = HSLtoRGB(hsl)
End Function

Public Function Darken(RGBColor As Long, Percent As Single)
'Darkens the color by a specifie percent, given as a Single

Dim hsl As HSLCol, L As Long
    If Percent <= 0 Then
        Darken = RGBColor
        Exit Function
    End If
    
    hsl = RGBtoHSL(RGBColor)
    L = hsl.Lum - (HSLMAX * Percent)
    If L < 0 Then L = 0
    hsl.Lum = L
    Darken = HSLtoRGB(hsl)
End Function

Public Function Blend(RGB1 As Long, RGB2 As Long, _
    Percent As Single) As Long
'This one doesn't really use the HSL routines, just the
'RGB Component routines.  I threw it in as a bonus ;)
'Takes two colors and blends them according to a
'percentage given as a Single
'For example, .3 will return a color 30% of the way
'between the first color and the second.
'.5, or 50%, will be an even blend (halfway)
'Can create some nice effects inside a For loop

Dim r As Integer, r1 As Integer, r2 As Integer, _
    g As Integer, g1 As Integer, g2 As Integer, _
    b As Integer, b1 As Integer, b2 As Integer
    
    If Percent >= 1 Then
        Blend = RGB2
        Exit Function
    ElseIf Percent <= 0 Then
        Blend = RGB1
        Exit Function
    End If
    
    r1 = RGBRed(RGB1)
    r2 = RGBRed(RGB2)
    g1 = RGBGreen(RGB1)
    g2 = RGBGreen(RGB2)
    b1 = RGBBlue(RGB1)
    b2 = RGBBlue(RGB2)
    
    r = ((r2 * Percent) + (r1 * (1 - Percent)))
    g = ((g2 * Percent) + (g1 * (1 - Percent)))
    b = ((b2 * Percent) + (b1 * (1 - Percent)))
    
    Blend = RGB(r, g, b)
End Function





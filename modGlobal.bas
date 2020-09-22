Attribute VB_Name = "modGlobal"
'--------------------------
' Scrollbar Designer Pro
'--------------------------
' Welcome to the sourcecode of Scrollbar Designer Pro.
' There's not alot of comments, I hope you find your way.
' Thanks for voting; http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=48809&lngWId=1
'
' / aDe
'
' http://www.ade.se

Global Const num3dLight = 0
Global Const numArrow = 1
Global Const numFace = 2
Global Const numShadow = 3
Global Const numDarkshadow = 4
Global Const numTrack = 5
Global Const numHighlight = 6

Global Const grpHSL = 1
Global Const grpRGB = 2
Global Const grpHEX = 3
Global Const grpHSLText = 4

'HSL/RGB Code
'Author: Andrew Gray
'Date: 9/10/2001 8:21:44 PM
'Link: http://abstractvb.com/code.asp?F=50&P=1&A=927
'Link: http://abstractvb.com/code.asp?F=50&P=1&A=926
Public Type HSL
    Hue As Integer
    Saturation As Integer
    Luminance As Integer
End Type

Public Type RGB
    Red As Integer
    Green As Integer
    Blue As Integer
End Type

Public Const HueMAX = 239, SatMAX = 240, LumMAX = 240

Function hexToRGB(sHex As String) As Long
        If Left$(sHex, 1) <> "#" Then sHex = "#" & sHex
       
        Dim R As Integer
        Dim G As Integer
        Dim B As Integer
        

        R = Val("&H" & Mid$(sHex, 2, 2)) ' red
        G = Val("&H" & Mid$(sHex, 4, 2)) ' green
        B = Val("&H" & Mid$(sHex, 6, 2)) ' blue

        hexToRGB = RGB(R, G, B)
End Function

Function MakeHex(Color) As String
  Dim R As Byte
  Dim G As Byte
  Dim B As Byte
  R = (Color And &HFF&)
  G = (Color And &HFF00&) / &H100&
  B = (Color And &HFF0000) / &H10000

  MakeHex = Right("0" & Hex(R), 2) & Right("0" & Hex(G), 2) & Right("0" & Hex(B), 2)
End Function


Public Function HSL(ByVal Hue As Integer, _
                         ByVal Saturation As Integer, _
                         ByVal Luminance As Integer) As Long
Dim RGBis As RGB
    RGBis = HSLtoRGB(Hue, Saturation, Luminance)
    If RGBis.Red < 0 Then RGBis.Red = 0
    If RGBis.Green < 0 Then RGBis.Green = 0
    If RGBis.Blue < 0 Then RGBis.Blue = 0
    HSL = RGB(RGBis.Red, RGBis.Green, RGBis.Blue)
End Function

Public Function HSLtoRGB(ByVal Hue As Integer, _
                         ByVal Saturation As Integer, _
                         ByVal Luminance As Integer) As RGB
'Author: Andrew Gray
'Date: 9/10/2001 8:18:23 PM
'Link: http://abstractvb.com/code.asp?F=50&P=1&A=926
'Modified by SB

    Dim pHue As Single
    Dim pSat As Single
    Dim pLum As Single
    Dim RetVal As RGB
    Dim pRed As Single
    Dim pGreen As Single
    Dim pBlue As Single
    Dim temp2 As Single
    Dim temp3() As Single
    Dim temp1 As Single
    Dim n As Integer

   ReDim temp3(0 To 2)
   
   pHue = Hue / HueMAX '239
   pSat = Saturation / SatMAX '239
   pLum = Luminance / LumMAX '239

   If pSat = 0 Then
      pRed = pLum!
      pGreen = pLum
      pBlue = pLum
   Else
      If pLum < 0.5 Then
         temp2 = pLum * (1 + pSat)
      Else
         temp2 = pLum + pSat - pLum * pSat
      End If
      temp1! = 2 * pLum! - temp2!
   
      temp3(0) = pHue + 1 / 3
      temp3(1) = pHue
      temp3(2) = pHue - 1 / 3
      
      For n = 0 To 2
         If temp3(n) < 0 Then temp3(n) = temp3(n) + 1
         If temp3(n) > 1 Then temp3(n) = temp3(n) - 1
      
         If 6 * temp3(n) < 1 Then
            temp3(n) = temp1 + (temp2 - temp1) * 6 * temp3(n)
         Else
            If 2 * temp3(n) < 1 Then
               temp3(n) = temp2
            Else
               If 3 * temp3(n%) < 2 Then
                  temp3(n%) = temp1 + (temp2 - temp1) _
                        * ((2 / 3) - temp3(n%)) * 6
               Else
                  temp3(n%) = temp1
                End If
             End If
          End If
       Next n%

       pRed = temp3(0)
       pGreen = temp3(1)
       pBlue = temp3(2)
    End If

    RetVal.Red = Int(pRed * 255)
    RetVal.Green = Int(pGreen * 255)
    RetVal.Blue = Int(pBlue * 255)
    
    HSLtoRGB = RetVal
End Function


Public Function RGBtoHSL(ByVal Red As Integer, _
                         ByVal Green As Integer, _
                         ByVal Blue As Integer) As HSL
'Author: Andrew Gray
'Date: 9/10/2001 8:21:44 PM
'Link: http://abstractvb.com/code.asp?F=50&P=1&A=927
'Modified by SB

    Dim pRed As Single
    Dim pGreen As Single
    Dim pBlue As Single
    Dim RetVal As HSL
    Dim pMax As Single
    Dim pMin As Single
    Dim pLum As Single
    Dim pSat As Single
    Dim pHue As Single
    
    pRed = Red / 255
    pGreen = Green / 255
    pBlue = Blue / 255
   
    If pRed > pGreen Then
       If pRed > pBlue Then
          pMax = pRed
       Else
          pMax = pBlue
       End If
    ElseIf pGreen > pBlue Then
        pMax = pGreen
    Else
        pMax = pBlue
    End If

    If pRed < pGreen Then
        If pRed < pBlue Then
            pMin = pRed
        Else
            pMin = pBlue
        End If
    ElseIf pGreen < pBlue Then
        pMin = pGreen
    Else
        pMin = pBlue
    End If

    pLum = (pMax + pMin) / 2
   
    If pMax = pMin Then
        pSat = 0
        pHue = 0
    Else
        If pLum < 0.5 Then
            pSat = (pMax - pMin) / (pMax + pMin)
        Else
            pSat = (pMax - pMin) / (2 - pMax - pMin)
        End If
        
        Select Case pMax!
            Case pRed
                pHue = (pGreen - pBlue) / (pMax - pMin)
            Case pGreen
                pHue = 2 + (pBlue - pRed) / (pMax - pMin)
            Case pBlue
                pHue = 4 + (pRed - pGreen) / (pMax - pMin)
        End Select
    End If

    RetVal.Hue = pHue * HueMAX \ 6
    If RetVal.Hue < 0 Then RetVal.Hue = RetVal.Hue + HueMAX + 1
    
    RetVal.Saturation = Int(pSat * SatMAX)
    RetVal.Luminance = Int(pLum * LumMAX)
    
    RGBtoHSL = RetVal
End Function


' Before/After-First/Last created/optimized by aDe
Function BeforeFirst(sIn, sFirst)
    If InStr(1, sIn, sFirst) Then
        BeforeFirst = Left(sIn, InStr(1, sIn, sFirst) - 1)
    Else
        BeforeFirst = ""
    End If
End Function

Function AfterFirst(sIn, sFirst)
    If InStr(1, sIn, sFirst) Then
        AfterFirst = Right(sIn, Len(sIn) - InStr(1, sIn, sFirst) - (Len(sFirst) - 1))
    Else
        AfterFirst = ""
    End If
End Function

Public Function AfterLast(sFrom, sAfterLast)
    If InStr(1, sFrom, sAfterLast) Then
        AfterLast = Right(sFrom, Len(sFrom) - InStrRev(sFrom, sAfterLast) - (Len(sAfterLast) - 1))
    Else
        AfterLast = ""
    End If
End Function

Public Function BeforeLast(sFrom, sBeforeLast)
    If InStr(1, sFrom, sBeforeLast) Then
        BeforeLast = Left(sFrom, InStrRev(sFrom, sBeforeLast) - 1)
    Else
        BeforeLast = ""
    End If
End Function

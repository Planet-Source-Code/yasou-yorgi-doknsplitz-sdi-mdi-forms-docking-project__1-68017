Attribute VB_Name = "modDefsGraphics"
Option Explicit
'--------------------------------------------------------
'Colors
'--------------------------------------------------------
Public Const COLOR_BLACK               As Long = vbBlack
Public Const COLOR_GREEN               As Long = vbGreen
Public Const COLOR_DARKGREEN           As Long = &HC000
Public Const COLOR_YELLOW              As Long = vbYellow
Public Const COLOR_RED                 As Long = vbRed
Public Const COLOR_BLUE                As Long = vbBlue
Public Const COLOR_WHITE               As Long = vbWhite
' //////////// Color constants. \\\\\\\\\\\\\\
Public Const COLOR_ACTIVEBORDER        As Long = 10
Public Const COLOR_ACTIVECAPTION       As Long = 2
Public Const COLOR_ADJ_MAX             As Long = 100
Public Const COLOR_ADJ_MIN             As Long = -100
Public Const COLOR_APPWORKSPACE        As Long = 12
Public Const COLOR_BACKGROUND          As Long = 1
Public Const COLOR_BTNFACE             As Long = 15
Public Const COLOR_BTNHIGHLIGHT        As Long = 20
Public Const COLOR_BTNLIGHT            As Long = 22
Public Const COLOR_BTNSHADOW           As Long = 16
Public Const COLOR_BTNTEXT             As Long = 18
Public Const COLOR_CAPTIONTEXT         As Long = 9
Public Const COLOR_GRAYTEXT            As Long = 17
Public Const COLOR_HIGHLIGHT           As Long = 13
Public Const COLOR_HIGHLIGHTTEXT       As Long = 14
Public Const COLOR_INACTIVEBORDER      As Long = 11
Public Const COLOR_INACTIVECAPTION     As Long = 3
Public Const COLOR_INACTIVECAPTIONTEXT As Long = 19
Public Const COLOR_MENU                As Long = 4
Public Const COLOR_MENUTEXT            As Long = 7
Public Const COLOR_SCROLLBAR           As Long = 0
Public Const COLOR_WINDOW              As Long = 5
Public Const COLOR_WINDOWFRAME         As Long = 6
Public Const COLOR_WINDOWTEXT          As Long = 8
Public Const NEWTRANSPARENT            As Long = 3 'used with SetBkMode()
Public Const WHITENESS                 As Long = &HFF0062
' //////////// Custom Colors \\\\\\\\\\\\\\\\\
Public Const vbMaroon                  As Long = 128
Public Const vbOlive                   As Long = 32896
Public Const vbNavy                    As Long = 8388608
Public Const vbPurple                  As Long = 8388736
Public Const vbTeal                    As Long = 8421376
Public Const vbGray                    As Long = 8421504
Public Const vbSilver                  As Long = 12632256
Public Const vbViolet                  As Long = 9445584
Public Const vbOrange                  As Long = 42495
Public Const vbGold                    As Long = 43724 '55295
Public Const vbIvory                   As Long = 15794175
Public Const vbPeach                   As Long = 12180223
Public Const vbTurquoise               As Long = 13749760
Public Const vbTan                     As Long = 9221330
Public Const vbBrown                   As Long = 17510
Public Const PATCOPY                   As Long = &HF00021
Public Const LF_FACESIZE               As Long = 32
'DrawText constants
Public Const DT_CALCRECT               As Long = &H400
Public Const DT_NOPREFIX               As Long = &H800
Public Const DT_SINGLELINE             As Long = &H20
Public Const DT_VCENTER                As Long = &H4
Public Const DT_END_ELLIPSIS           As Long = &H8000&
Public Const DT_MODIFYSTRING           As Long = &H10000
Public Const DT_LEFT                   As Long = &H0
Public Const DT_CENTER                 As Long = &H1
Public Const DT_BOTTOM                 As Long = &H8
Public Const DT_RIGHT                  As Long = &H2
Public Const DT_TITLEBAR               As Long = DT_NOPREFIX Or DT_SINGLELINE Or DT_LEFT Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
'Borders
Public Const BF_FLAT                   As Long = &H4000
Public Const BF_MONO                   As Long = &H8000
Public Const BF_SOFT                   As Long = &H1000 ' For softer buttons
'--- Types Declaration
Public Type POINTAPI
   X                                   As Long
   Y As Long
End Type
Public Type RECT
   Left                                As Long
   Top                                 As Long
   Right                               As Long
   Bottom                              As Long
End Type
Public Type LOGFONT
   lfHeight                            As Long
   lfWidth                             As Long
   lfEscapement                        As Long
   lfOrientation                       As Long
   lfWeight                            As Long
   lfItalic                            As Byte    '0=false; 255=true
   lfUnderline                         As Byte    '0=f; 255=t
   lfStrikeOut                         As Byte    '0=f; 255=t
   lfCharSet                           As Byte
   lfOutPrecision                      As Byte
   lfClipPrecision                     As Byte
   lfQuality                           As Byte
   lfPitchAndFamily                    As Byte
   lfFaceName(32)                      As Byte
End Type
Public Enum GradientDirectionCts
   [gdHorizontal] = 0
   [gdVertical] = 1
   [gdDownwardDiagonal] = 2
   [gdUpwardDiagonal] = 3
End Enum
Public Const PS_SOLID                  As Integer = 0
Public Const BLACK_PEN                 As Integer = 7
Public Const DIB_RGB_COLORS            As Long = 0
Public Const LOGPIXELSX                As Long = 88 '  Logical pixels/inch in X
Public Const LOGPIXELSY                As Long = 90 '  Logical pixels/inch in Y
Public Const FW_NORMAL                 As Long = 400
Public Const FW_BOLD                   As Long = 700
Public Const FF_DONTCARE               As Long = 0
Public Const DEFAULT_PITCH             As Long = 0
Public Const DEFAULT_CHARSET           As Long = 1
Public Const CLR_INVALID               As Long = -1
Public Const GM_ADVANCED               As Long = 2
' lfQuality Constants:
Public Const DEFAULT_QUALITY           As Long = 0 ' Appearance of the font is set to default
Public Const DRAFT_QUALITY             As Long = 1 ' Appearance is less important that PROOF_QUALITY.
Public Const PROOF_QUALITY             As Long = 2 ' Best character quality
Public Const NONANTIALIASED_QUALITY    As Long = 3 ' Don't smooth font edges even if system is set to smooth font edges
Public Const ANTIALIASED_QUALITY       As Long = 4 ' Ensure font edges are smoothed if system is set to smooth font edges
Public Type BITMAPINFOHEADER
   biSize                              As Long
   biWidth                             As Long
   biHeight                            As Long
   biPlanes                            As Integer
   biBitCount                          As Integer
   biCompression                       As Long
   biSizeImage                         As Long
   biXPelsPerMeter                     As Long
   biYPelsPerMeter                     As Long
   biClrUsed                           As Long
   biClrImportant                      As Long
End Type
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function TextOutA Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function SetGraphicsMode Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Public Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
'*****************************************************
Public Function BlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR, Optional ByVal alpha As Long = 128) As Long
'*****************************************************
   Dim lCFrom              As Long
   Dim lCTo                As Long
   Dim lDstB               As Long
   Dim lDstG               As Long
   Dim lDstR               As Long
   Dim lSrcB               As Long
   Dim lSrcG               As Long
   Dim lSrcR               As Long
10   lCFrom = TranslateColor(oColorFrom)
20   lCTo = TranslateColor(oColorTo)
30   lSrcR = lCFrom And &HFF
40   lSrcG = (lCFrom And &HFF00&) \ &H100&
50   lSrcB = (lCFrom And &HFF0000) \ &H10000
60   lDstR = lCTo And &HFF
70   lDstG = (lCTo And &HFF00&) \ &H100&
80   lDstB = (lCTo And &HFF0000) \ &H10000
90   BlendColor = RGB(((lSrcR * alpha) / 255) + ((lDstR * (255 - alpha)) / 255), ((lSrcG * alpha) / 255) + ((lDstG * (255 - alpha)) / 255), ((lSrcB * alpha) / 255) + ((lDstB * (255 - alpha)) / 255))
End Function
'*****************************************************
Public Sub PaintGradient(ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal GradientDirection As GradientDirectionCts)
'*****************************************************
   '-- A minor check
   Dim B1                  As Long
   Dim B2                  As Long
   Dim dB                  As Long
   Dim dG                  As Long
   Dim dR                  As Long
   Dim G1                  As Long
   Dim G2                  As Long
   Dim i                   As Long
   Dim iEnd                As Long
   Dim iGrad               As Long
   Dim iOffset             As Long
   Dim j                   As Long
   Dim jEnd                As Long
   Dim lBits()             As Long
   Dim lGrad()             As Long
   Dim R1                  As Long
   Dim R2                  As Long
   Dim Scan                As Long
   Dim uBIH                As BITMAPINFOHEADER
10   If (Width < 1 Or Height < 1) Then Exit Sub
     '-- Decompose colors
20   Color1 = Color1 And &HFFFFFF
30   R1 = Color1 Mod &H100&
40   Color1 = Color1 \ &H100&
50   G1 = Color1 Mod &H100&
60   Color1 = Color1 \ &H100&
70   B1 = Color1 Mod &H100&
80   Color2 = Color2 And &HFFFFFF
90   R2 = Color2 Mod &H100&
100   Color2 = Color2 \ &H100&
110   G2 = Color2 Mod &H100&
120   Color2 = Color2 \ &H100&
130   B2 = Color2 Mod &H100&
      '-- Get color distances
140   dR = R2 - R1
150   dG = G2 - G1
160   dB = B2 - B1
      '-- Size gradient-colors array
170   Select Case GradientDirection
         Case [gdHorizontal]
180         ReDim lGrad(0 To Width - 1)
190      Case [gdVertical]
200         ReDim lGrad(0 To Height - 1)
210      Case Else
220         ReDim lGrad(0 To Width + Height - 2)
230      End Select
      '-- Calculate gradient-colors
240   iEnd = UBound(lGrad())
250   If (iEnd = 0) Then
         '-- Special case (1-pixel wide gradient)
260      lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
270   Else
280      For i = 0 To iEnd
290         lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
300         Next i
310      End If
      '-- Size DIB array
320   ReDim lBits(Width * Height - 1) As Long
330   iEnd = Width - 1
340   jEnd = Height - 1
350   Scan = Width
      '-- Render gradient DIB
360   Select Case GradientDirection
         Case [gdHorizontal]
370         For j = 0 To jEnd
380            For i = iOffset To iEnd + iOffset
390               lBits(i) = lGrad(i - iOffset)
400               Next i
410            iOffset = iOffset + Scan
420            Next j
430      Case [gdVertical]
440         For j = jEnd To 0 Step -1
450            For i = iOffset To iEnd + iOffset
460               lBits(i) = lGrad(j)
470               Next i
480            iOffset = iOffset + Scan
490            Next j
500      Case [gdDownwardDiagonal]
510         iOffset = jEnd * Scan
520         For j = 1 To jEnd + 1
530            For i = iOffset To iEnd + iOffset
540               lBits(i) = lGrad(iGrad)
550               iGrad = iGrad + 1
560               Next i
570            iOffset = iOffset - Scan
580            iGrad = j
590            Next j
600      Case [gdUpwardDiagonal]
610         iOffset = 0
620         For j = 1 To jEnd + 1
630            For i = iOffset To iEnd + iOffset
640               lBits(i) = lGrad(iGrad)
650               iGrad = iGrad + 1
660               Next i
670            iOffset = iOffset + Scan
680            iGrad = j
690            Next j
700      End Select
      '-- Define DIB header
710   With uBIH
720      .biSize = 40
730      .biPlanes = 1
740      .biBitCount = 32
750      .biWidth = Width
760      .biHeight = Height
770      End With
      '-- Paint it!
780   Call StretchDIBits(hdc, xLeft, yTop, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)
End Sub
'*****************************************************
Public Sub pOLEFontToLogFont(fntThis As StdFont, hdc As Long, tLF As LOGFONT)
'*****************************************************
   ' Convert an OLE StdFont to a LOGFONT structure:
   Dim iChar               As Integer
   Dim sFont               As String
10   With tLF
20      sFont = fntThis.Name
        ' There is a quicker way involving StrConv and CopyMemory, but
        ' this is simpler!:
30      For iChar = 1 To Len(sFont)
40         .lfFaceName(iChar - 1) = CByte(Asc(mId$(sFont, iChar, 1)))
50         Next iChar
        ' Based on the Win32SDK documentation:
60      .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hdc, LOGPIXELSY)), 72)
70      .lfItalic = fntThis.Italic
80      If (fntThis.Bold) Then
90         .lfWeight = FW_BOLD
100      Else
110         .lfWeight = FW_NORMAL
120         End If
130      .lfUnderline = fntThis.Underline
140      .lfStrikeOut = fntThis.Strikethrough
150      .lfCharSet = fntThis.Charset
160      .lfQuality = ANTIALIASED_QUALITY
170      End With
End Sub
'*****************************************************
Public Sub SetRedraw(hWnd As Long, b As Boolean)
'*****************************************************
10   Call SendMessageAsLong(hWnd, WM_SETREDRAW, IIf(b, 1, 0), 0)
End Sub
'*****************************************************
Public Function TranslateColor(ByVal clr As OLE_COLOR, Optional hpal As Long = 0) As Long
'*****************************************************
   ' Routine       : TranslateColor
   ' Created by    : Marclei V Silva
   ' Date-Time     : 02/10/0010:20:19
   ' Inputs        :
   ' Outputs       :
   ' Credits       : Extracted from VB KB Article
   ' Modifications :
   ' Purpose   : Converts an OLE_COLOR to RGB color
10   If OleTranslateColor(clr, hpal, TranslateColor) Then
20      TranslateColor = CLR_INVALID
30      End If
End Function

' Yorgi's 4Matz [Feb 28,2007 23:58:50] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

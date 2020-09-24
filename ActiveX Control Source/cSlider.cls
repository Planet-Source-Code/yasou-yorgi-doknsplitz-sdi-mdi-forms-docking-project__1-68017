VERSION 1.0 CLASS
BEGIN
MultiUse = -1                                     'True
Persistable = 0                                   'NotPersistable
DataBindingBehavior = 0                           'vbNone
DataSourceBehavior  = 0                           'vbNone
MTSTransactionMode  = 0                           'NotAnMTSObject
END
Attribute VB_Name = "clsSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
   X                                   As Long
   Y As Long
End Type
Private Type RECT
   Left                                As Long
   Top                                 As Long
   Right                               As Long
   Bottom                              As Long
End Type
Private Type BITMAP                               '24 bytes
   bmType                              As Long
   bmWidth                             As Long
   bmHeight                            As Long
   bmWidthBytes                        As Long
   bmPlanes                            As Integer
   bmBitsPixel                         As Integer
   bmBits                              As Long
End Type
Private Const IDC_SIZENS               As Long = 32645&
Private Const IDC_SIZEWE               As Long = 32644&
Private Const IDC_NO                   As Long = 32648&
Private Const R2_NOTXORPEN             As Long = 10 '  DPxn
Private Const PATINVERT                As Long = &H5A0049 ' (DWORD) dest = pattern XOR dest
Private Const DSTINVERT                As Long = &H550009 ' (DWORD) dest = (NOT dest)
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub ClipCursorRect Lib "user32" Alias "ClipCursor" (lpRect As RECT)
Private Declare Sub ClipCursorClear Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long)
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function LoadCursorLong Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Enum ESPLTOrientationConstants
   cSPLTOrientationHorizontal = 1
   cSPLTOrientationVertical = 2
End Enum
Public Enum ESPLTPanelConstants
   cSPLTLeftOrTopPanel = 1
   cSPLTRightOrBottomPanel = 2
End Enum
Private lSliderPos                     As Long
Private mlngDelta                      As Long
Private mlngSliderSize                 As Long
Private mblnInDrag                     As Boolean
Private m_tPInitial                    As POINTAPI
Private mlngSplitInitial               As Long
Private m_hBrush                       As Long
Private mlngPattern(0 To 3)            As Long
Private m_tSlider                      As RECT
Private m_hCursor                      As Long
Private m_eOrientation                 As ESPLTOrientationConstants
Private tSlideArea                     As RECT
'*****************************************************
Private Sub Class_Initialize()
'*****************************************************
   Dim i                   As Long
10   m_eOrientation = cSPLTOrientationHorizontal
20   m_hCursor = LoadCursorLong(0, IDC_SIZENS)
30   mlngSliderSize = 4
40   lSliderPos = 0
50   For i = 0 To 3
60      mlngPattern(i) = &HAAAA5555
70      Next i
End Sub
'*****************************************************
Private Sub Class_Terminate()
'*****************************************************
10   DestroyBrush
20   If Not (m_hCursor = 0) Then
30      DestroyCursor m_hCursor
40      End If
End Sub
'*****************************************************
Private Function CreateBrush() As Boolean
'*****************************************************
   Dim hBm                 As Long
   Dim tbm                 As BITMAP
10   DestroyBrush
     ' Create a monochrome bitmap containing the desired pattern:
20   tbm.bmType = 0
30   tbm.bmWidth = 16
40   tbm.bmHeight = 8
50   tbm.bmWidthBytes = 2
60   tbm.bmPlanes = 1
70   tbm.bmBitsPixel = 1
80   tbm.bmBits = VarPtr(mlngPattern(0))
90   hBm = CreateBitmapIndirect(tbm)
     ' Make a brush from the bitmap bits
100   m_hBrush = CreatePatternBrush(hBm)
      '// Delete the useless bitmap
110   DeleteObject hBm
End Function
'*****************************************************
Public Property Get Delta() As Long
'*****************************************************
10   Delta = mlngDelta
End Property
'*****************************************************
Private Sub DestroyBrush()
'*****************************************************
10   If Not (m_hBrush = 0) Then
20      DeleteObject m_hBrush
30      m_hBrush = 0
40      End If
End Sub
'*****************************************************
Friend Sub MouseDown(Button As Integer, lTop As Long, lLeft As Long, lBottom As Long, lRight As Long)
'*****************************************************
   Dim tP                  As POINTAPI
10   If (Button = vbLeftButton) Then
20      mblnInDrag = True
30      CreateBrush
40      GetCursorPos tP
50      LSet m_tPInitial = tP
60      mlngSplitInitial = lSliderPos
70      tSlideArea.Top = lTop
80      tSlideArea.Left = lLeft
90      tSlideArea.Bottom = lBottom
100      tSlideArea.Right = lRight
110      ClipCursorRect tSlideArea
120      If (m_eOrientation = cSPLTOrientationVertical) Then
130         m_tSlider.Left = tSlideArea.Left + lSliderPos
140         m_tSlider.Right = m_tSlider.Left + mlngSliderSize
150         m_tSlider.Top = tSlideArea.Top
160         m_tSlider.Bottom = tSlideArea.Bottom
170      Else
180         m_tSlider.Left = tSlideArea.Left
190         m_tSlider.Right = tSlideArea.Right
200         m_tSlider.Top = tSlideArea.Top + lSliderPos
210         m_tSlider.Bottom = m_tSlider.Top + mlngSliderSize
220         End If
230      pDrawSlider
240      End If
End Sub
'*****************************************************
Public Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
   Dim tP                  As POINTAPI
10   If mblnInDrag Then
20      SetCursor m_hCursor
30      GetCursorPos tP
40      pDrawSlider
50      If (m_eOrientation = cSPLTOrientationVertical) Then
60         lSliderPos = mlngSplitInitial + (tP.X - m_tPInitial.X)
70      Else
80         lSliderPos = mlngSplitInitial + (tP.Y - m_tPInitial.Y)
90         End If
100      If (m_eOrientation = cSPLTOrientationVertical) Then
110         m_tSlider.Left = tSlideArea.Left + lSliderPos
120         m_tSlider.Right = m_tSlider.Left + mlngSliderSize
130         m_tSlider.Top = tSlideArea.Top
140         m_tSlider.Bottom = tSlideArea.Bottom
150      Else
160         m_tSlider.Left = tSlideArea.Left
170         m_tSlider.Right = tSlideArea.Right
180         m_tSlider.Top = tSlideArea.Top + lSliderPos
190         m_tSlider.Bottom = m_tSlider.Top + mlngSliderSize
200         End If
210      pDrawSlider
220      End If
End Sub
'*****************************************************
Public Sub MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
   Dim tP                  As POINTAPI
10   If (mblnInDrag) Then
20      ClipCursorClear 0&
30      GetCursorPos tP
40      pDrawSlider
50      If (m_eOrientation = cSPLTOrientationVertical) Then
60         mlngDelta = tP.X - m_tPInitial.X
70      Else
80         mlngDelta = tP.Y - m_tPInitial.Y
90         End If
100      lSliderPos = mlngSplitInitial + mlngDelta
110      If lSliderPos < 0 Then lSliderPos = 0
120      DestroyBrush
130      mblnInDrag = False
140      End If
End Sub
'*****************************************************
Public Property Get Orientation() As ESPLTOrientationConstants
'*****************************************************
10   Orientation = m_eOrientation
End Property
'*****************************************************
Public Property Let Orientation(ByVal eOrientation As ESPLTOrientationConstants)
'*****************************************************
10   If Not (m_eOrientation = eOrientation) Then
20      m_eOrientation = eOrientation
30      If Not (m_hCursor = 0) Then
40         DestroyCursor m_hCursor
50         End If
60      If (m_eOrientation = cSPLTOrientationHorizontal) Then
70         m_hCursor = LoadCursorLong(0, IDC_SIZENS)
80      Else
90         m_hCursor = LoadCursorLong(0, IDC_SIZEWE)
100         End If
110      End If
End Property
'*****************************************************
Private Sub pDrawSlider()
'*****************************************************
   Dim hOldBrush           As Long
   Dim lhDC                As Long
10   lhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
20   hOldBrush = SelectObject(lhDC, m_hBrush)
30   PatBlt lhDC, m_tSlider.Left, m_tSlider.Top, m_tSlider.Right - m_tSlider.Left, m_tSlider.Bottom - m_tSlider.Top, PATINVERT
40   SelectObject lhDC, hOldBrush
50   DeleteDC lhDC
End Sub
'*****************************************************
Public Property Get Position() As Long
'*****************************************************
10   Position = lSliderPos
End Property
'*****************************************************
Public Property Let Position(ByVal lPosition As Long)
'*****************************************************
10   If (lPosition <> lSliderPos) Then
20      lSliderPos = lPosition
30      End If
End Property
'*****************************************************
Public Property Let SliderSize(ByVal lSize As Long)
'*****************************************************
10   If Not (mlngSliderSize = lSize) Then
20      If (lSize < 0) Then
30         Err.Raise 380, App.EXEName & ".clsSlider"
40      Else
50         mlngSliderSize = lSize
60         End If
70      End If
End Property
'*****************************************************
Public Property Get SliderSize() As Long
'*****************************************************
10   SliderSize = mlngSliderSize
End Property

' Yorgi's 4Matz [Feb 28,2007 23:58:51] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

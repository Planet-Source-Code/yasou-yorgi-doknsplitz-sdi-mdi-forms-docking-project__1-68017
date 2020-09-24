VERSION 5.00
Begin VB.UserControl ctlTitleBar
Appearance      =   0                             'Flat
AutoRedraw      =   -1                            'True
BackColor       =   &H80000002&
ClientHeight    =   360
ClientLeft      =   0
ClientTop       =   0
ClientWidth     =   4770
FillColor       =   &H00FFFFC0&
BeginProperty Font
Name            =   "Arial"
Size            =   9.75
Charset         =   0
Weight          =   700
Underline       =   0                             'False
Italic          =   0                             'False
Strikethrough   =   0                             'False
EndProperty
ForeColor       =   &H8000000E&
ScaleHeight     =   24
ScaleMode       =   3                             'Pixel
ScaleWidth      =   318
Begin DoknSplitz.ctlButton ctlCloseBtn
Height          =   255
Left            =   4410
TabIndex        =   0
TabStop         =   0                             'False
Top             =   30
Width           =   255
_ExtentX        =   450
_ExtentY        =   450
End
End
Attribute VB_Name = "ctlTitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** File Name   : ctlTitleBar.ctl                                             **
'** Description : A custom title bar ActiveX control                          **
'** Usage       : To provide interface for the user to move the control in    **
'**               ControlManager at run-time                                  **
'** Dependencies: ctlButton, mdlAPI                                           **
'** Members     :                                                             **
'**   * Collections: -                                                        **
'**   * Objects    : -                                                        **
'**   * Property   : CloseVisible                                             **
'**   * Methods    : -                                                        **
'**   * Events     : Click, CloseClick, DblClick, MouseDown, MouseMove,       **
'**                  MouseUp, Move, MoveBegin, MoveEnd                        **
'** Modifications:                                                            **
'** 11/14/03 Theo Z - Last modified                                           **
'** 12/03/06 Yorgi  - Added support for captions and gradients                **
'*******************************************************************************
Option Explicit
'--- Property Variables
Private lLenCaption                    As Long
Private m_oStartColor                  As Long
Private m_oEndColor                    As Long
Private Ret                            As Long
Private hFont                          As Long
Private PrevFont                       As Long
Private rctBarStripe                   As RECT
Private m_Orientation                  As TBarOrientation
Private m_sCaption                     As String  'Yorgi: add Caption heading
Private mblnCloseVisible               As Boolean
Private m_TBarType                     As TBarTypes
'--- Private Constants
Private Const conSpacerBarThickness    As Long = 4
Private Const constlBarTop             As Long = 4
Private Const mconSpacerGap            As Long = 5 'gap width in pixels between the title bar and its control button
Private Const mconBorderGap            As Long = 1 'gap width in pixels around the border
'--- Private Variables
Private eGradDir                       As GradientDirectionCts 'gradient direction
Private lBarStripeLength               As Long    'length of barstripes
Private rctCaption                     As RECT    'rect for the caption
Private rctCalcCaptionSize             As RECT    'rect for the caption
Private mblnDrag                       As Boolean 'indicating whether the user is dragging the title bar
Private mposLastDrag                   As POINTAPI 'last (x,y) coordinate of the dragging action
Private lngCloseButtonWidth            As Long    'close button width if visible else zero
'Description- Occurs when no initial caption is present.  Gives proggy a change
'             to dynamically set from owner form
Public Event Caption()
'Description- Occurs when the user presses and then realeses a mouse button over
'             the control
Public Event Click()
'Description- Occurs when the user presses and then realeses a mouse button and
'             then presses and releases it again over the control
Public Event DblClick()
'Description- Occurs when the user presses and then releases a mouse button over
'             the close button
Public Event CloseClick()
'Description- Occurs when the user presses a mouse button over the control
'Arguments  - Button, Shift, X, Y (see reference for MouseDown event in MSDN for
'                                  the description of the arguments)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'Description- Occurs when the user moves the mouse over the control
'Arguments  - Button, Shift, X, Y (see reference for MouseMove event in MSDN for
'                                  the description of the arguments)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'Description- Occurs when the user releases a mouse button over the control
'Arguments  - Button, Shift, X, Y (see reference for MouseUp event in MSDN for
'                                  the description of the arguments)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'Description- Occurs when the user moves the mouse after the BeginMove event and
'             before the EndMove event
'Arguments  - Shift (an integer that corresponds to the state of the SHIFT,
'                    CTRL, and ALT keys)
'           - bHitchHiker is a doknForm object hitching a ride in the move process
Public Event Move(ByVal Shift As Integer, bHitchHiker As Boolean)
'Description- Occurs when the user presses a mouse left-button over the custom-
'             drawing title bar
'Arguments  - Shift (an integer that corresponds to the state of the SHIFT,
'                    CTRL, and ALT keys)
'           - bHitchHiker is a doknForm object hitching a ride in the move process
Public Event MoveBegin(ByVal Shift As Integer, bHitchHiker As Boolean)
'Description- Occurs when the user releases the mouse button after the BeginMove
'             event
'Arguments  - Shift (an integer that corresponds to the state of the SHIFT,
'                    CTRL, and ALT keys)
Public Event MoveEnd(ByVal Shift As Integer, ByRef dfHitchhiker As DokNForm, ByRef blnSuccess As Boolean)
'*****************************************************
Public Property Let Caption(ByRef sCaption As String)
'*****************************************************
10   m_sCaption = sCaption
20   lLenCaption = Len(m_sCaption)
30   If lLenCaption Then
        'calculate how much room text takes up
40      DrawText hdc, m_sCaption, lLenCaption, rctCalcCaptionSize, DT_CALCRECT
50      End If
60   Call UserControl.PropertyChanged("Caption")
End Property
'*****************************************************
Public Property Get Caption() As String
'*****************************************************
10   Caption = m_sCaption
End Property
'*****************************************************
Public Property Let CloseVisible(ByRef blnCloseVisible As Boolean)
'*****************************************************
   ' Purpose    - Sets a value that determines whether a close button in the title
   '              bar is visible
   ' Input      - blnCloseVisible (the new CloseVisible property value)
10   mblnCloseVisible = blnCloseVisible
20   ctlCloseBtn.Visible = mblnCloseVisible
30   If mblnCloseVisible Then
40      lngCloseButtonWidth = 14
50   Else
60      lngCloseButtonWidth = 0
70      End If
80   UserControl_Resize
End Property
'*****************************************************
Public Property Get CloseVisible() As Boolean
'*****************************************************
   ' Purpose    - Returns a value that determines whether a close button in the
   '              title bar is visible
10   CloseVisible = mblnCloseVisible
End Property
'*****************************************************
Private Sub ctlCloseBtn_Click()
'*****************************************************
   ' Purpose    - Raises custom event CloseClick
   ' Effect     - As specified
10   RaiseEvent CloseClick
End Sub
'*****************************************************
Private Sub DrawTitleBar()
'*****************************************************
   ' Purpose    - Draws the title bar
   ''''    Dim rctTemp As RECT
   'draw the background gradient
   Dim captionFont         As LOGFONT
   Dim hbr                 As Long
10   UserControl.Cls                              'cleanup previous painting which restores background color
20   If m_oStartColor <> m_oEndColor Then
30      If m_Orientation = TBO_VERTICAL Then
40         PaintGradient hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_oEndColor, m_oStartColor, eGradDir
50      Else
60         PaintGradient hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_oStartColor, m_oEndColor, eGradDir
70         End If
80      End If
90   If lLenCaption Then
        'create rectangle dimensions separator bars
100      If m_Orientation = TBO_VERTICAL Then
110         pOLEFontToLogFont UserControl.Font, hdc, captionFont
120         captionFont.lfEscapement = 900
130         hFont = CreateFontIndirect(captionFont)
140         PrevFont = SelectObject(hdc, hFont)
150         End If
         'Draw the caption text
160      DrawText hdc, m_sCaption, lLenCaption, rctCaption, DT_SINGLELINE 'Or DT_END_ELLIPSIS
170      If PrevFont Then                         'if exists, restore the previous Font
180         SelectObject hdc, PrevFont
190         PrevFont = 0
200         End If
210      End If
      'make sure we have room to draw separator bars
220   If lBarStripeLength > 0 Then
230      hbr = CreateSolidBrush(TranslateColor(vbButtonFace))
240      FillRect hdc, rctBarStripe, hbr
250      DeleteObject hbr
260      DrawEdge UserControl.hdc, rctBarStripe, mdlAPI.EDGE_RAISED, mdlAPI.BF_RECT
270      If (m_TBarType And TBT_DOUBLESTRIPE) = TBT_DOUBLESTRIPE Then
280         If m_Orientation = TBO_VERTICAL Then
290            rctBarStripe.Left = rctBarStripe.Left + conSpacerBarThickness + 1
300            rctBarStripe.Right = rctBarStripe.Left + conSpacerBarThickness
310         Else
320            rctBarStripe.Top = rctBarStripe.Top + conSpacerBarThickness + 1
330            rctBarStripe.Bottom = rctBarStripe.Top + conSpacerBarThickness
340            End If
350         DrawEdge UserControl.hdc, rctBarStripe, mdlAPI.EDGE_RAISED, mdlAPI.BF_RECT
360         End If
370      End If
End Sub
'*****************************************************
Public Property Let EndColor(oColor As Long)
'*****************************************************
10   If (m_oEndColor <> oColor) Then
20      m_oEndColor = oColor
30      DrawTitleBar
40      End If
End Property
'*****************************************************
Public Property Get EndColor() As Long
'*****************************************************
10   EndColor = m_oEndColor
End Property
'*****************************************************
Public Property Let Orientation(ByVal eOrientation As TBarOrientation)
'*****************************************************
10   m_Orientation = eOrientation
20   If m_Orientation = TBO_HORIZONTAL Then
30      eGradDir = gdHorizontal                   'set gradient direction
40   Else
50      eGradDir = gdVertical
60      End If
70   Call UserControl.PropertyChanged("Orientation")
End Property
'*****************************************************
Public Property Get Orientation() As TBarOrientation
'*****************************************************
10   Orientation = m_Orientation
End Property
'*****************************************************
Public Property Let StartColor(ByVal oColor As Long)
'*****************************************************
10   If (m_oStartColor <> oColor) Then
20      m_oStartColor = oColor
30      UserControl.BackColor = m_oStartColor
40      DrawTitleBar
50      End If
End Property
'*****************************************************
Public Property Get StartColor() As Long
'*****************************************************
10   StartColor = m_oStartColor
End Property
'*****************************************************
Public Property Get TBarType() As TBarTypes
'*****************************************************
   ' Purpose    - Returns the title bar draw type
10   TBarType = m_TBarType
End Property
'*****************************************************
Public Property Let TBarType(blnTBarType As TBarTypes)
'*****************************************************
   ' Purpose    - Sets  the title bar draw type
   ' Input      - blnTBarType (the new TBarType property value)
10   m_TBarType = blnTBarType
End Property
'*****************************************************
Private Sub UserControl_Click()
'*****************************************************
   ' Purpose    - Raises custom event Click
10   RaiseEvent Click
End Sub
'*****************************************************
Private Sub UserControl_DblClick()
'*****************************************************
10   RaiseEvent DblClick
End Sub
'*****************************************************
Private Sub UserControl_Initialize()
'*****************************************************
   'set scalemode to pixels and avoid extra repetative math conversions during drawtitle
10   UserControl.ScaleMode = vbPixels
     'init gradient colors
20   m_oStartColor = UserControl.BackColor
     '''''30  m_oEndColor = &H808000
30   m_oEndColor = &HFFFFC0
End Sub
'*****************************************************
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
   ' Purpose    - Begins move the title bar virtually and raises custom event MouseDown
   ' Effect     - * If Button = vbLeft Button, as specified
   '              * Otherwise no effect
   ' Inputs     - Button, Shift, X, Y
10   RaiseEvent MouseDown(Button, Shift, X, Y)
20   If Button = vbLeftButton Then
30      mblnDrag = True
40      mposLastDrag.X = X
50      mposLastDrag.Y = Y
60      RaiseEvent MoveBegin(Shift, False)
70      End If
End Sub
'*****************************************************
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
   ' Purpose    - Moves the title bar virtually and raises custom event Move
   ' Assumption - The UserControl_MouseDown procedure has been called
   ' Effect     - * If mblnDrag = true, as specified
   '              * Otherwise, no effect
   ' Inputs     - Index, Button, Shift, x, y
10   If mblnDrag Then
20      RaiseEvent Move(Shift, 0)
30      mposLastDrag.X = X
40      mposLastDrag.Y = Y
50   Else
60      RaiseEvent MouseMove(Button, Shift, X, Y)
70      End If
End Sub
'*****************************************************
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
   ' Purpose    - Ends the title bar virtual move action and raises custom event Moved
   ' Effect     - * If mblnDrag = true, as specified
   '              * Otherwise, no effect
   ' Inputs     - Index, Button, Shift, x, y
   Dim blnSuccess          As Boolean
10   RaiseEvent MouseUp(Button, Shift, X, Y)
20   If mblnDrag Then
30      mblnDrag = False
40      RaiseEvent MoveEnd(Shift, Nothing, blnSuccess)
50      End If
End Sub
'*****************************************************
Private Sub UserControl_Resize()
'*****************************************************
   ' Purpose    - Adjusts the components inside the control to agree with the control's size
   Dim lHeight             As Long
   Dim lLeft               As Long
   Dim lTop                As Long
   Dim lWidth              As Long
10   With rctBarStripe
20      If m_Orientation = TBO_VERTICAL Then
30         If lngCloseButtonWidth Then
              'init the Bar Stripe vertical rect with Close Button
40            .Left = constlBarTop
50            .Right = .Left + conSpacerBarThickness
              'move the ctlCloseBtn on right edge and redraw control
60            With ctlCloseBtn
70               lWidth = lngCloseButtonWidth
80               lHeight = lWidth
90               lLeft = mconBorderGap
100               .Move lLeft, mconBorderGap, lWidth, lHeight
110               End With
120            .Top = lngCloseButtonWidth + mconSpacerGap
130            .Bottom = UserControl.ScaleHeight - .Top - mconSpacerGap
140         Else
               'init the Bar Stripe vertical rect without Close Button
150            .Top = mconSpacerGap
160            .Bottom = UserControl.ScaleHeight - mconSpacerGap
170            End If
180         If lLenCaption Then
190            With rctCalcCaptionSize
                  ' add a leading gap area
200               rctCaption.Top = UserControl.ScaleHeight - mconSpacerGap
210               rctCaption.Bottom = rctCaption.Top - .Right
220               rctCaption.Right = UserControl.ScaleWidth - mconBorderGap
230               rctCaption.Left = mconBorderGap
240               End With
               'adjust Bar Stripe vertical rect for Caption
250            .Bottom = rctCaption.Bottom - mconSpacerGap
260            End If
            'calculate the Bar Stripe Length
270         lBarStripeLength = .Bottom - .Top
280      Else
            'init the Bar Stripe horizontal rect with Close Button
290         .Top = constlBarTop
300         .Bottom = .Top + conSpacerBarThickness
310         If lngCloseButtonWidth Then
               'move the ctlCloseBtn on top edge
320            With ctlCloseBtn
330               lWidth = lngCloseButtonWidth
340               lHeight = lWidth
350               lLeft = UserControl.ScaleWidth - lWidth - mconBorderGap
360               .Move lLeft, mconBorderGap, lWidth, lHeight
370               End With
380            .Left = mconSpacerGap
390            .Right = lLeft - mconSpacerGap
400         Else
               'init the Bar Stripe horizontal rect without Close Button
410            .Left = mconSpacerGap
420            .Right = UserControl.ScaleWidth - mconSpacerGap
430            End If
440         If lLenCaption Then
450            rctCaption = rctCalcCaptionSize
460            rctCaption.Left = rctCalcCaptionSize.Left + mconSpacerGap
470            rctCaption.Right = rctCaption.Right + mconSpacerGap
               'adjust Bar Stripe vertical rect for Caption
480            .Left = rctCaption.Right + mconSpacerGap
490            End If
            'calculate the Bar Stripe Length
500         lBarStripeLength = .Right - .Left
510         End If
520      End With
530   DrawTitleBar
End Sub

' Yorgi's 4Matz [Feb 28,2007 23:58:50] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

VERSION 5.00
Begin VB.UserControl ControlManager
Alignable       =   -1                            'True
Appearance      =   0                             'Flat
AutoRedraw      =   -1                            'True
BackColor       =   &H80000005&
ClientHeight    =   1500
ClientLeft      =   0
ClientTop       =   0
ClientWidth     =   1455
ControlContainer=   -1                            'True
EditAtDesignTime=   -1                            'True
FillColor       =   &H00404040&
PaletteMode     =   0                             'Halftone
ScaleHeight     =   1500
ScaleWidth      =   1455
ToolboxBitmap   =   "ctlControlManager.ctx":0000
Begin DoknSplitz.ctlRect crecControl
Height          =   825
Left            =   300
TabIndex        =   3
Top             =   450
Visible         =   0                             'False
Width           =   855
_ExtentX        =   1508
_ExtentY        =   1455
End
Begin DoknSplitz.ctlTitleBar ctbTitlebar
Height          =   315
Index           =   0
Left            =   300
TabIndex        =   2
Top             =   60
Visible         =   0                             'False
Width           =   855
_ExtentX        =   1508
_ExtentY        =   556
End
Begin VB.PictureBox picSlider
Appearance      =   0                             'Flat
BackColor       =   &H80000005&
BorderStyle     =   0                             'None
FillColor       =   &H0080C0FF&
FillStyle       =   0                             'Solid
ForeColor       =   &H80000008&
Height          =   900
Left            =   1230
MousePointer    =   3                             'I-Beam
ScaleHeight     =   60
ScaleMode       =   3                             'Pixel
ScaleWidth      =   5
TabIndex        =   1
TabStop         =   0                             'False
Top             =   210
Visible         =   0                             'False
Width           =   75
End
Begin VB.PictureBox picSplitter
Appearance      =   0                             'Flat
BackColor       =   &H80000005&
BorderStyle     =   0                             'None
ClipControls    =   0                             'False
FillColor       =   &H0080C0FF&
FillStyle       =   0                             'Solid
ForeColor       =   &H80000008&
Height          =   1380
Index           =   9999
Left            =   0
ScaleHeight     =   1380
ScaleWidth      =   120
TabIndex        =   0
TabStop         =   0                             'False
Top             =   0
Visible         =   0                             'False
Width           =   120
End
End
Attribute VB_Name = "ControlManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An ActiveX control to allow the user to resize docked controls at run time"
'*******************************************************************************
'** File Name   - DoknSplitz                                                  **
'** Author      - Yorgi (yorgi@omnisoftsystems.com)                           **
'** Description - A supersized version of TheoZ's VB Control Manager.  Now    **
'**               resize, move and show/hide of design time controls &        **
'**               runtime forms docking capability for SDI & MDI projects.    **
'** Credits     - * Special thanks to Theo Zacharias for VB Control Manager   **
'**               * To Steve McMahon excellent code www.vbaccelerator.com     **
'**               * and all the great contributions found on PSC & DevX       **
'*******************************************************************************
'******* H I S T O R Y   L O G *************************************************
'*******************************************************************************
'** 11/14/03 v1.0.5  TheoZ- ControlManager.ctl Last modified
'** 11/05/06 v1.1.0  Yorgi- Applied code reformatting, sort, and numbered functions/subs, ERL error handling
'** 11/07/06 v1.1.1  Yorgi- Convert to Paul_Caton.Subclass w/ minor err handling mods
'** 11/23/06 v1.1.17 Yorgi- Added Runtime control capability
'** 12/04/06 v1.1.21 Yorgi- Added IDoknForm interface
'** 12/10/06 v1.1.25 Yorgi- Integrating DockedForms functionality
'** 12/15/07 v1.1.31 Yorgi- Support draggable UnDocked forms as Hitchhikers
'** 12/19/07 v1.1.37 Yorgi- Various performance changes & code cleanup, every nanosec helps to fight VB-Bloat
'** 01/02/07 v1.1.39 Yorgi- Added BatchBuild (small speed enhancement) to consolidate rebuild and painting overhead
'** 01/08/07 v1.1.42 Yorgi- Split ControlManager functionality into VCtrlManager,VCntrlRebuildAll,VCntrlAdd,,VCntrlRemove
'** 01/11/07 v1.1.43 Yorgi- Add TitleBar draw vertical capability, apply default actions (doublebars=dockable, singlebar=moveable)
'** 01/19/07 v1.1.52 Yorgi- GetDropTarget valid return only if cursor is inside Usercontrol
'** 01/25/07 v1.1.57 Yorgi- Add event notify for FormAdd & FormRemove
'** 01/28/07 v1.1.59 Yorgi- Modified the Activate logic to handle aligned DoknSplitz control
'** 01/29/07 v1.1.60 Yorgi- Added a Slider Bar to size aligned DoknSplitz controls
'** 02/03/07 v1.1.61 Yorgi- Modified TitleBar handling to support TBarCreate & TBarRemove functions
'** 02/04/07 v1.1.62 Yorgi- Refresh is now re-entrant, allows for "rebuild later" on errors (ie Usercontrol too small to draw componenets)
'** 02/15/07 v1.1.64 Yorgi- Added error handling to Stretch and checked for valid Splitter.IdCtlFriends
'** 02/20/07 v1.1.67 Yorgi- Overhauled all classes (except clsId) to use "ControlKey" as the collection Key instead of Index.
'** 02/22/07 v1.1.68 Yorgi- Fix(iSubclass_WndProc):Don't call Activate on SIZE_MINIMIZE it's MADNESS
'**   Why? Because the ContainedCntrls collection uses a push to the BOTTOM method of indexing anytime the Visible property
'**   of an embedded control is changed.  This renders the ContainedCntrls index useless as a reference in other stored variables/collections.
'**   When docking a form for example, the VCtrlManager had to always re-Index all of the referencing collections to stay in sync.
'** 02/24/07 v1.1.71 Yorgi- Fix(MDI Slider):Inital alignment did not reposition VirtualControls.Left properly for Slider
'** 02/26/07 v1.1.72 Yorgi- SDI/MDI Demos: Replaced RichTextBox controls with InternetControl to support 4Matz generated html docs
'*******************************************************************************
'****** T H I N G S  T O  D O **************************************************
'*******************************************************************************
'** - Moving ctlRect leaves a terrible trailing rect effect.  Must find a way another way! Maybe BitBlt an image???
'** - Always looking for additional functionality, but mostly performance gains!!!!
'*******************************************************************************
Option Explicit
Private Const mconModuleName           As String = "ControlManager"
Private Const mconHostCtlPrefixName    As String = "HC"
Public Enum genmMoveDestination
   mdControlTop
   mdControlRight
   mdControlBottom
   mdControlLeft
   mdEdgeTop
   mdEdgeRight
   mdEdgeBottom
   mdEdgeLeft
   mdSplitter
End Enum
Private Type typUCInnerDimensions
   Top                                 As Long
   Left                                As Long
   Width                               As Long
   Height                              As Long
End Type
'--- Collection Variables- Used to represents virtual controls and splitters
Private WithEvents mVirtualControls       As clsControls
Attribute mVirtualControls.VB_VarHelpID = -1
Private WithEvents mSplitters             As clsSplitters
Attribute mSplitters.VB_VarHelpID = -1
'--- Property Variables
Private mblnFillContainer              As Boolean
Private mlngMarginBottom               As Long
Private mlngMarginLeft                 As Long
Private mlngMarginRight                As Long
Private mlngMarginTop                  As Long
Private mblnUnloadFrmOnClose           As Boolean 'Unload docked form OnClose or just hide
Private mlngAppHosthWnd                As Long
'--- PropBag Names
Private Const mconUnloadFrmOnClose     As String = "UnloadFrmOnClose"
Private Const mconActiveColor          As String = "ActiveColor"
Private Const mconBackColor            As String = "BackColor"
Private Const mconClipCursor           As String = "ClipCursor"
Private Const mconEnable               As String = "Enable"
Private Const mconFillContainer        As String = "FillContainer"
Private Const mconLiveUpdate           As String = "LiveUpdate"
Private Const mconMarginBottom         As String = "MarginBottom"
Private Const mconMarginLeft           As String = "MarginLeft"
Private Const mconMarginRight          As String = "MarginRight"
Private Const mconMarginTop            As String = "MarginTop"
Private Const mconSize                 As String = "Size"
Private Const mconTitleBar_CloseVisible As String = "TitleBar_CloseVisible"
Private Const mconTitleBar_Height      As String = "TitleBar_Height"
Private Const mconTitleBar_Visible     As String = "TitleBar_Visible"
Private Const mconTitleBar_TBarType    As String = "TitleBar_TBarType"
Private Const mconTitleBar_Position    As String = "TitleBar_Position"
'--- Property Default Values
Private Const mconDefaultUnloadFrmOnClose As Boolean = False
Private Const mconDefaultFillContainer As Boolean = True
Private Const mconDefaultMarginBottom  As Long = 0
Private Const mconDefaultMarginLeft    As Long = 0
Private Const mconDefaultMarginRight   As Long = 0
Private Const mconDefaultMarginTop     As Long = 0
'--- Saved Procedure-level variables lost in subclassing process for MouseDown, MouseMove, and MouseUp events
Private mScIndex                       As Integer
Private mScButton                      As Integer
Private mScShift                       As Integer
Private mScX                           As Single
Private mScY                           As Single
'--- Other Variables
Private mblnLastRefreshOK              As Boolean 'Refresh fails if not enough room to expand all splitters so rebuild
Private mlngSliderThickness            As Long
Private mblnDragSplitter               As Boolean 'indicating whether the user is dragging the splitter
Private mblnControlMoved               As Boolean 'indicating whether a control has just been moved by the user
Private mblnSplitterMoved              As Boolean 'indicating whether a splitter has just been moved by the user
Private mblnVisibleSave                As Boolean 'to restore the Visible property of the control's instance
Private mlngHwndParent                 As Long    'the handle of the control's container
Private mlngHwndRoot                   As Long    'the handle of the root window of the control
'the x- or y- coordinate (depends on the active Splitter 's orientation)
' where the user strats to drag it
Private mlngDragStart                  As Long
'previous mouse pointer coordinate relative to the splitter (note- this variable is
' used to make sure the custom event MouseMove works properly)
Private muposPrev                      As POINTAPI
'allows adding multiple controls before processed by VCtrlManager & painting functions
Private mblnBatchBuild                 As Boolean
Private oSub                           As cSubclass
Private oDockedForms                   As DokNForms
Private mblnRefreshInProgress          As Boolean
Private oSlider                        As clsSlider
Attribute oSlider.VB_VarHelpID = -1
Private mtypUCInside                   As typUCInnerDimensions
Private mrctSlideArea                  As RECT    ' for aligned control slider area
Private mrctUserControl                As RECT
'--- Implements the Interface
Implements IDoknForm                              'sends WinEvents from DoknForms to UserControl
Implements iSubclass                              'All Hail the Great and Powerful Subclasser!
Implements TitleBar
'-------------------------------
' ActiveX Control Custom Events
'-------------------------------
'Description- Occurs when a control has just been closed by the user
'Arguments  - IdControl (a value that uniquely identifies the control that has
'                        just been closed by the user)
Public Event ControlAfterClose(ByVal sIdControl As String)
'Description- Occurs after the user presses a close button of certain control
'             and before the control is closed
'Arguments  - * IdControl (a value that uniquely identifies the control that
'                          about to be closed)
'             * Cancel (setting this argument to true stops the control from
'                       closing)
Public Event ControlBeforeClose(ByVal sIdControl As String, ByRef Cancel As Boolean)
'Description- Occurs when the user is moving a control
'Arguments  - * IdControl (a value that uniquely identifies the control that is
'                          being moved by the user)
'             * Shift (an integer that corresponds to the state of the SHIFT,
'                      CTRL, and ALT keys)
'             * Left (an integer indicating the x-coordinate for the left edge
'                     of the current position of the rectangle that represents
'                     the moving control)
'             * Top (an integer indicating the y-coordinate for the top edge of
'                    the current position of the rectangle that represents the
'                    moving control)
'             * Width (an integer indicating the current width of the rectangle
'                      that represents the moving control)
'             * Height (an integer indicating the current height of the
'                       rectangle that represents the moving control)
Public Event ControlMove(ByVal sIdControl As String, ByVal Shift As Integer, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
'Description- Occurs when the user is about to move a control, i.e. the first
'             time the rectangle that represents the moving control occurs
'Arguments  - * IdControl (a value that uniquely identifies the control that is
'                          ready to be moved)
'             * Shift (an integer that corresponds to the state of the SHIFT,
'                      CTRL, and ALT keys)
Public Event ControlMoveBegin(ByVal sIdControl As String, ByVal Shift As Integer)
'Description- Occurs when the user is finished moving a control, i.e. when the
'             rectangle that represents the moving control disappears
'Arguments  - * IdControl (a value that uniquely identifies the control that
'                          has just been moved)
'             * Shift (an integer that corresponds to the state of the SHIFT,
'                      CTRL, and ALT keys)
'             * Moved (a value that determines whether the control is moved)
Public Event ControlMoveEnd(ByVal sIdControl As String, ByVal Shift As Integer, ByVal Moved As Boolean)
'Description- Occurs when the user presses and then realeses a mouse button over
'             a splitter
'Arguments  - IdSplitter (a value that uniquely identifies the splitter that has
'                         just been clicked by the user)
Public Event SplitterClick(ByVal IdSplitter As Long)
Attribute SplitterClick.VB_Description = "Occurs when the user presses and then realeses a mouse button over a splitter"
'Description- Occurs when the user presses and then realeses a mouse button and
'             then presses and releases it again over a splitter
'Arguments  - IdSplitter (a value that uniquely identifies the splitter that has
'                         just been double-clicked by the user)
Public Event SplitterDblClick(ByVal IdSplitter As Long)
Attribute SplitterDblClick.VB_Description = "Occurs when the user presses and then realeses a mouse button and then presses and releases it again over a splitter"
'Description- Occurs when the user presses a mouse button over a splitter
'Arguments  - * IdSplitter (a value that uniquely identifies the splitter where
'                           the user presses a mouse button over)
'             * Button, Shift, X, Y (see reference for MouseDown event in
'                                    MSDN Library for the description of the
'                                    arguments)
Public Event SplitterMouseDown(ByVal IdSplitter As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute SplitterMouseDown.VB_Description = "Occurs when the user presses a mouse button over a splitter"
'Description- Occurs when the user moves a mouse over a splitter without moving
'             the splitter
'Arguments  - * IdSplitter (a value that uniquely identifies a splitter where
'                           the user moves a mouse over)
'             * Button, Shift, X, Y (see reference for MouseMove event in
'                                    MSDN Library for the description of the
'                                    arguments)
Public Event SplitterMouseMove(ByVal IdSplitter As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute SplitterMouseMove.VB_Description = "Occurs when the user moves a mouse over a splitter without moving the splitter"
'Description- Occurs when the user releases a mouse button over a splitter
'             without previously moving the splitter
'Arguments  - * IdSplitter (a value that uniquely identifies the splitter where
'                           the user releases a mouse button over)
'             * Button, Shift, X, Y (see reference for MouseUp event in
'                                    MSDN for the description of the arguments)
Public Event SplitterMouseUp(ByVal IdSplitter As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute SplitterMouseUp.VB_Description = "Occurs when the user releases a mouse button over a splitter without previously moving the splitter"
'Description- Occurs when the user is moving a splitter
'Arguments  - * IdSplitter (a value that uniquely identifies the splitter that
'                           is being moved by the user)
'             * Shift, X, Y (see reference for MouseMove event in MSDN for the
'                            description of the arguments)
Public Event SplitterMove(ByVal IdSplitter As Long, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute SplitterMove.VB_Description = "Occurs when the user is moving a splitter"
'Description- Occurs when the user is about to move a splitter
'Arguments  - * IdSplitter (A value that uniquely identifies the splitter that
'                           is about to be moved by the user)
'             * Shift, X, Y (see reference for MouseDown event in MSDN for the
'                            description of the arguments)
Public Event SplitterMoveBegin(ByVal IdSplitter As Long, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute SplitterMoveBegin.VB_Description = "Occurs when the user is about to move a splitter"
'Description- Occurs when the user is finished moving a splitter
'Arguments  - * IdSplitter (a value that uniquely identifies the splitter that
'                           has just been moved by the user)
'             * Shift, X, Y (see reference for MouseUp event in MSDN for the
'                            description of the arguments)
Public Event SplitterMoveEnd(ByVal IdSplitter As Long, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute SplitterMoveEnd.VB_Description = "Occurs when the user presses and then realeses a mouse button over a control title bar"
'Description- Occurs when the user presses and then realeses a mouse button over
'             a control title bar
'Arguments  - IdControl (a value that uniquely identifies the control whose
'                  title bar has just been clicked by the user)
Public Event TitleBarClick(ByVal sIdControl As String)
'Description- Occurs when the user presses and then realeses a mouse button and
'             then presses and releases it again over a control title bar
'Arguments  - IdControl (a value that uniquely identifies the control that owns the title bar)
Public Event TitleBarDblClick(ByVal sIdControl As String)
'Description- Occurs when the user presses a mouse button over a control title bar
'Arguments  - * IdControl (a value that uniquely identifies the control that own
'                  the title bar where the user presses a mouse button over)
'             * Button, Shift, X, Y (see reference for MouseDown event in
'                  MSDN Library for the description of the arguments)
Public Event TitleBarMouseDown(ByVal sIdControl As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'Description- Occurs when the user moves the mouse over a control title bar
'             without moving the control
'Arguments  - * IdControl (A value that uniquely identifies the control that own
'                  the title bar where the user moves a mouse over)
'             * Button, Shift, X, Y (see reference for MouseMove event in
'                  MSDN Library for the description of the arguments)
Public Event TitleBarMouseMove(ByVal sIdControl As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'Description- Occurs when the user releases a mouse button over a control title
'             bar without previously moving the control
'Arguments  - * IdControl (a value that uniquely identifies the control that own
'                  the title bar where the user releases a mouse button over)
'             * Button, Shift, X, Y (see reference for MouseUp event in
'                  MSDN Library for the description of the arguments)
Public Event TitleBarMouseUp(ByVal sIdControl As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'-----Use 4Matz Events Section to only doc our changes--------------
Public Event FormAddComplete(ByVal sFormName As String, ByVal sKey As String, ByVal sIdControl As String) '4Matz:New
'Description- Occurs after successful completion of FormAdd
'Arguments  - sFormName (Name of form)
'             sKey (a key value that uniquely identifies the control)
'             IdControl (a long value that uniquely identifies the control)
'** 01/25/07 Yorgi- Add event notify for FormAdd & FormRemove
Public Event FormRemoveComplete(ByVal sFormName As String) '4Matz:New
'Description- Occurs when the user releases a mouse button over a control title
'             bar without previously moving the control
'Arguments  - * IdControl (a value that uniquely identifies the control that own
'                  the title bar where the user releases a mouse button over)
'             * sCaption (the controls new caption)
'** 01/25/07 Yorgi- Add event notify for FormAdd & FormRemove
Public Event TitleBarCaption(ByVal sIdControl As String, ByRef sCaption As String) '4Matz:New
'Description- Occurs after successful completion of FormRemove
'Arguments  - sFormName (Name of form)
'** 01/11/07 v1.1.43 Yorgi- Add TitleBar caption for contained controls
'*****************************************************
Private Sub Activate() '4Matz:Changed
Attribute Activate.VB_Description = "Activates and resize the control to meet its container size with respect to the control's margin property and FillContainer property"
'*****************************************************
   ' Purpose    - Activate and resize the control to meet its container size with
   '              respect to the control's margin property and FillContainer property
   ' Assumption - The parent of the control has ScaleWidth and ScaleHeight property
   ' Note       - This is the main method of the control. This method should be
   '              called whenever its container is loaded. Also this method should
   '              be called everytime its container's size is changed so that the
   '              FillContainer property would work.
   '** 01/28/07 Yorgi- ReWork the Activate logic to handle aligned controls
   Dim lngHeight           As Long               'the new height of the control
   Dim lngWidth            As Long               'the new width of the control
10   On Error GoTo Activate_Err
     'TraceCtl  "Activate Extender.Parent.hWnd:" & Extender.Parent.hWnd
20   If mlngHwndRoot = 0 Then
30      mlngHwndRoot = mdlAPI.GetAncestor(Extender.Container.hWnd, mdlAPI.GA_ROOT)
40      oSub.Subclass mlngHwndRoot, Me
50      On Error Resume Next                      'property may not be available
60      InitSlider Extender.Align                 'if aligned need to init the slider
70      On Error GoTo 0
80      End If
90   If mblnFillContainer Then
        '''''** 01/28/07 Yorgi- ReCalc dimensions to handle aligned controls
        ''''' Unfortunately, UserControl_Resize will be called again
100      With UserControl.Parent
110         Select Case Extender.Align
               Case vbAlignNone
120               lngWidth = .ScaleWidth - mlngMarginRight - mlngMarginLeft
130               lngHeight = .ScaleHeight - mlngMarginBottom - mlngMarginTop
140               If lngWidth < 0 Then lngWidth = 0
150               If lngHeight < 0 Then lngHeight = 0
                  'TraceCtl  "Activate Extender.Move Left:" & mlngMarginLeft & ", Top:" & mlngMarginTop & ", Width:" & lngWidth & ", Height:" & lngHeight
160               Extender.Move mlngMarginLeft, mlngMarginTop, lngWidth, lngHeight
170            Case vbAlignLeft, vbAlignRight
180               lngWidth = .Width - mlngMarginRight - mlngMarginLeft
190               If lngWidth > 0 Then
                     'TraceCtl  "Activate Extender.Width:" & lngWidth
200                  Extender.Width = lngWidth
210                  End If
220            Case vbAlignTop, vbAlignBottom
230               lngHeight = .Height - mlngMarginTop - mlngMarginBottom
240               If lngHeight > 0 Then
                     'TraceCtl  "Activate Extender.Height:" & lngHeight
250                  Extender.Height = lngHeight
260                  End If
270            End Select
280         End With
290   Else
300      UserControl_Resize
310      End If
      'TraceCtl  "Activate UserControl Left:" & UserControl.ScaleLeft & ", Top:" & UserControl.ScaleTop & ", Width:" & UserControl.ScaleWidth & ", Height:" & UserControl.ScaleHeight
320   Activate_Exit:
330   On Error GoTo 0
340   Exit Sub
350   Activate_Err:
360   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", Activate", mconModuleName
370   Resume Activate_Exit
End Sub
'*****************************************************
Public Property Get ActiveColor() As OLE_COLOR
Attribute ActiveColor.VB_Description = "Returns/sets the background color used to display the splitter when the user moves it in none live update mode"
'*****************************************************
   ' Purpose    - Returns the background color used to display a splitter when the
   '              user moves it in none live update mode
10   ActiveColor = mSplitters.ActiveColor
End Property
'*****************************************************
Public Property Let ActiveColor(lngActiveColor As OLE_COLOR)
'*****************************************************
   ' Purpose    - Sets the background color used to display a splitter when the
   '              user moves it in none live update mode
   ' Input      - lngActiveColor (the new ActiveColor property value)
10   mSplitters.ActiveColor = lngActiveColor
20   PropertyChanged mconActiveColor
End Property
'*****************************************************
Private Function AdjustedHeight(ctl As Control, octl As clsControl) As Long   '4Matz:Changed
Attribute AdjustedHeight.VB_Description = "Returns the adjusted height of control ctl"
'*****************************************************
   ' Purpose    - Returns the adjusted height of control ctl
   ' Inputs     - * ctl
   '              * octl (the virtual control of control ctl)
   ' Note       - This function is used to avoid flickering effect in LiveUpdate
   '              mode for list box control or other controls that inherit it
   Dim lngHeightFactor     As Long               'the height of each item in the list box
10   If Not (TypeOf ctl Is ListBox) Then
20      If Not (TypeOf ctl Is DirListBox) Then
30         If Not (TypeOf ctl Is FileListBox) Then
40            AdjustedHeight = octl.Height
50            Exit Function
60            End If
70         End If
80      End If
90   lngHeightFactor = mdlAPI.SendMessage(ctl.hWnd, mdlAPI.LB_GETITEMHEIGHT, 0&, 0&) * Screen.TwipsPerPixelY
100   AdjustedHeight = (((octl.Height - octl.MinHeight) \ lngHeightFactor) * lngHeightFactor) + octl.MinHeight
End Function
'*****************************************************
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display all splitters"
'*****************************************************
   ' Purpose    - Returns the background color used to display all splitters
10   BackColor = mSplitters.BackColor
End Property
'*****************************************************
Public Property Let BackColor(lngBackColor As OLE_COLOR)
'*****************************************************
   ' Purpose    - Sets the background color used to display all splitters
   ' Input      - lngBackColor (the new BackColor property value)
10   mSplitters.BackColor = lngBackColor
20   PropertyChanged mconBackColor
30   If Ambient.UserMode = False Then Refresh     'only during design time
End Property
'*****************************************************
Public Property Let BatchBuild(bBatch As Boolean) '4Matz:New
'*****************************************************
   ' Purpose    - Performance gain by batching multiple control adds/mods before
   '** Build Manager processing or painting functions occur.  When false, Build Manager is
   '** automatically called to process all pending changes.
   '** WARNING- VCtlIdxs are re-assigned during calls to VCtrlRebuildAll so do not store locally.  Always
   '** retrieve from the DockedForm or Controls object (ie df.VCtlIdx)
   '** 01/20/07 Yorgi- Create BatchBuild speed enhancement to consolidate rebuild and painting overhead
10   If Not (mblnBatchBuild = bBatch) Then
20      mblnBatchBuild = bBatch
30      If Not bBatch Then VCtrlManager
40      End If
End Property
'*****************************************************
Public Property Get BatchBuild() As Boolean '4Matz:New
'*****************************************************
   ' Purpose    - Create BatchBuild speed enhancement to consolidate rebuild and painting overhead
10   BatchBuild = mblnBatchBuild
End Property
'*****************************************************
Public Property Let ClipCursor(blnClipCursor As Boolean)
'*****************************************************
   ' Purpose    - Sets a value that determines whether the mouse pointer is
   '              confined to the virtual splitter minimum and maximum x- and
   '              y-coordinate when the user moves a splitter
   ' Input      - blnClipCursor (the new ClipCursor property value)
10   mSplitters.ClipCursor = blnClipCursor
20   PropertyChanged mconClipCursor
End Property
'*****************************************************
Public Property Get ClipCursor() As Boolean
Attribute ClipCursor.VB_Description = "Returns/sets a value that determines whether the mouse pointer is confined to the virtual splitter minimum and maximum x- and y-coordinate when the user moves a splitter"
'*****************************************************
   ' Purpose    - Returns a value that determines whether the mouse pointer is
   '              confined to the virtual splitter minimum and maximum x- and
   '              y-coordinate when the user moves a splitter
10   ClipCursor = mSplitters.ClipCursor
End Property
'*****************************************************
Public Property Get Controls() As clsControls
Attribute Controls.VB_Description = "Returns a collection whose elements represent each virtual control in a Control Manager object"
'*****************************************************
10   Set Controls = mVirtualControls
End Property
'*****************************************************
Private Sub CreateSplitr(lngIdx As Long) '4Matz:New
'*****************************************************
   ' Purpose    - Loads a picSplitter object
10   On Error Resume Next
20   Load picSplitter(lngIdx)                     '-- Creates the new PictureBox control instances to represent the splitter
30   On Error GoTo 0
     'TraceCtl  "..CreateSplitr:" & lngIdx
40   picSplitter(lngIdx).Visible = True
End Sub
'*****************************************************
Private Sub ctbTitleBar_Click(Index As Integer)
'*****************************************************
   ' Purpose    - Raises custom event TitleBarClick
   ' Input      - Index
10   RaiseEvent TitleBarClick(ctbTitlebar(Index).Tag)
End Sub
'*****************************************************
Private Sub ctbTitleBar_CloseClick(Index As Integer) '4Matz:Changed
'*****************************************************
   ' Purpose    - Closes the control at run-time, re-arranges the other controls
   '              and raises ControlBeforeClose and ControlAfterClose event
   ' Effect     - See the codes
   ' Input      - Index (the id of the control which will be closed)
   Dim blnCancel           As Boolean
   Dim df                  As DokNForm
   Dim sOwner              As String
10   sOwner = ctbTitlebar(Index).Tag
20   RaiseEvent ControlBeforeClose(sOwner, blnCancel)
30   If Not blnCancel Then
40      If mblnUnloadFrmOnClose Then
50         If VCtrlIdxToDoknForm(sOwner, df) Then 'look for a docked form object
60            blnCancel = FormRemove(df.DockedForm)
70            Set df = Nothing
80            End If
90         End If
100      If Not blnCancel Then
110         ShowControl sOwner, False
120         RaiseEvent ControlAfterClose(sOwner)
130         End If
140      End If
End Sub
'*****************************************************
Private Sub ctbTitleBar_DblClick(Index As Integer) '4Matz:Changed
'*****************************************************
   ' Purpose    - Raises custom event TitleBarDblClick
   ' Input      - Index
   '** 01/11/07 Yorgi: If allowed to float, then use this event to also undock control
   Dim df                  As DokNForm
   Dim sOwner              As String
10   sOwner = ctbTitlebar(Index).Tag
20   RaiseEvent TitleBarDblClick(Index)
30   If VCtrlIdxToDoknForm(sOwner, df) Then       'look for a docked form object
        'cant click on TBar unless docked, duh, but we check the state anyway
40      If df.State = DS_Docked Then
50         If (df.Style And DSFloat) = DSFloat Then
60            UnDock df
70            End If
80         End If
90      Set df = Nothing
100      End If
End Sub
'*****************************************************
Private Sub ctbTitlebar_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'*****************************************************
   ' Purpose    - Raises custom event TitleBarMouseDown
   ' Inputs     - Index, Button, Shift, x, y
10   muposPrev.X = X
20   muposPrev.Y = Y
30   RaiseEvent TitleBarMouseDown(Index, Button, Shift, X, Y)
End Sub
'*****************************************************
Private Sub ctbTitlebar_MouseMove(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'*****************************************************
   ' Purpose    - Raises custom event TitleBarMouseMove
   ' Inputs     - Index, Button, Shift, x, y
10   If Not mblnControlMoved Then
20      If ((X <> muposPrev.X) Or (Y <> muposPrev.Y)) Then RaiseEvent TitleBarMouseMove(Index, Button, Shift, X, Y)
30      End If
End Sub
'*****************************************************
Private Sub ctbTitlebar_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'*****************************************************
   ' Purpose    - Raises custom event TitleBarMouseUp
   ' Inputs     - Index, Button, Shift, x, y
10   If Not mblnControlMoved Then RaiseEvent TitleBarMouseUp(Index, Button, Shift, X, Y)
End Sub
'*****************************************************
Private Sub ctbTitleBar_Move(Index As Integer, ByVal Shift As Integer, bHitchHiker As Boolean) '4Matz:Changed
'*****************************************************
   ' Purpose    - Moves the control at run-time and raises ControlMove event
   ' Effects    - * If the cursor is on a splitter which doesn't belong to the
   '                control or the cursor is on the edge of the ControlManager
   '                control, then the drop guider rectangle has been shown
   '              * Otherwise, the guider rectangle position has been adjusted
   '                based on the cursor position
   ' Inputs     - * Index (the id of the control which will be moved)
   '              * Shift (an integer that corresponds to the state of the SHIFT,
   '                       CTRL, and ALT keys)
   '** 01/15/07 Yorgi - added bHitchHiker flag so we can use the same logic for all moving objects
   '                    regardless of runtime(docked/undocked) or design time controls
   Dim blnShowDropRect     As Boolean            'T/F Show drop guider rectangle
   Dim IdSpl               As Long
   Dim sIdCtlName          As String
   Dim sOwner              As String
   Dim udeTargetType       As genmMoveDestination
   Dim uposCursor          As POINTAPI           'cursor position relative to UserControl
   Dim urecDrop            As RECT               'drop guider rect size and position
10   mblnControlMoved = True                      'guilty until proven innocent
     'There has to be an api to compare sizes between 2 rect, do it this way for now..yuk
20   sOwner = ctbTitlebar(Index).Tag
30   With mVirtualControls(sOwner)
40      If crecControl.Left = .Left Then
50         If crecControl.Top = .Top Then
60            If crecControl.Width = .Width Then
70               If crecControl.Height = .Height Then
80                  mblnControlMoved = False
90                  End If
100               End If
110            End If
120         End If
130      End With
140   GetDropTarget blnShowDropRect, udeTargetType, sIdCtlName, IdSpl, uposCursor
150   blnShowDropRect = blnShowDropRect And Not IsRectNearSource(sOwner, sIdCtlName, IdSpl, udeTargetType)
160   If blnShowDropRect Then
170      urecDrop = GetDropRect(sOwner, sIdCtlName, IdSpl, udeTargetType)
         '-- urecDrop.Left = gconUninitializedLong means that the drop guider rectangle's
         '   size is bigger than the minimum size of the control
180      blnShowDropRect = urecDrop.Left <> gconUninitializedLong
190      End If
200   If blnShowDropRect Then
         '-- Show the drop guider rectangle
210      crecControl.Move urecDrop.Left, urecDrop.Top, urecDrop.Right - urecDrop.Left, urecDrop.Bottom - urecDrop.Top
220      If bHitchHiker Then crecControl.Visible = True ' possible drop zone so show the guider
230   Else
240      If bHitchHiker Then crecControl.Visible = False ' not a drop zone so don't show the guider
         '-- Update the guider rectangle position based on the cursor position
250      crecControl.UpdatePosition
260      End If
270   If mblnControlMoved Then RaiseEvent ControlMove(Index, Shift, crecControl.Left, crecControl.Top, crecControl.Width, crecControl.Height)
End Sub
'*****************************************************
Private Sub ctbTitleBar_MoveBegin(Index As Integer, ByVal Shift As Integer, bHitchHiker As Boolean) '4Matz:Changed
'*****************************************************
   ' Purpose    - Initializes all things needed to move the control at run-time
   ' Effect     - The guider rectangle has been shown
   ' Inputs     - * Index (the id of the control which will be moved)
   '              * Shift (an integer that corresponds to the state of the SHIFT,
   '                       CTRL, and ALT keys)
   ' This subclassing below is used to handle the possibility of the user
   '   swithing to another application while dragging the splitter
10   mScIndex = Index
20   mScShift = Shift
30   If Not bHitchHiker Then
40      oSub.AddMsg mlngHwndRoot, WM_ACTIVATE, MSG_AFTER
50   Else
60      Screen.MousePointer = vbCrosshair
70      End If
80   With mVirtualControls(ctbTitlebar(Index).Tag)
90      crecControl.Move .Left, .Top, .Width, .Height
100      End With
110   crecControl.ZOrder
120   crecControl.Visible = True
130   RaiseEvent ControlMoveBegin(Index, Shift)
End Sub
'*****************************************************
Private Sub ctbTitleBar_MoveEnd(Index As Integer, ByVal Shift As Integer, ByRef dfHitchhiker As DokNForm, blnSuccess As Boolean) '4Matz:Changed
'*****************************************************
   ' Purpose    - Ends the run-time control move action
   ' Effect     - * The guider rectangle has been hidden
   '              * If the drop target is valid, the control has been moved and
   '                the other controls position and size have been re-arranged
   ' Inputs     - * Index (the id of the control which will be moved)
   '              * Shift (an integer that corresponds to the state of the SHIFT, CTRL, and ALT keys)
   ' Variables for GetDropTarget parameters
   '** 01/21/07 Yorgi- Support draggable UnDocked form as Hitchhiker
   Dim blnTargetValid      As Boolean
   Dim IdSpl               As Long
   Dim sIdCtl              As String
   Dim sSrc                As String
   Dim udeTargetType       As genmMoveDestination
   Dim uposCursor          As POINTAPI
10   sSrc = ctbTitlebar(Index).Tag
20   If dfHitchhiker Is Nothing Then
30      oSub.DelMsg mlngHwndRoot, WM_ACTIVATE, MSG_AFTER
40   Else
50      Screen.MousePointer = vbDefault
60      End If
70   If crecControl.Visible Then
80      GetDropTarget blnTargetValid, udeTargetType, sIdCtl, IdSpl, uposCursor
90      If blnTargetValid Then blnTargetValid = blnTargetValid And Not IsRectNearSource(sSrc, sIdCtl, IdSpl, udeTargetType)
100      If blnTargetValid Then
            'now we have a valid drop target
110         If Not dfHitchhiker Is Nothing Then
120            Dock dfHitchhiker, True            '-- HitchHiker has to be closed in order to get here
130            End If
140         blnSuccess = MoveControl(sSrc, udeTargetType, sIdCtl, IdSpl)
150      Else
160         blnSuccess = False
170         crecControl.Visible = False
180         End If
190      blnSuccess = blnTargetValid And blnSuccess 'indicating whether the move action is succesful
200      If mblnControlMoved Then
210         RaiseEvent ControlMoveEnd(sSrc, Shift, blnSuccess)
220         End If
230      End If
240   mblnControlMoved = False
End Sub
'*****************************************************
Public Sub DetachAll() '4Matz:New
'*****************************************************
   ' Purpose   - Stop all dockedform subclassing.  Used when Main App is unloading
   Dim df                  As DokNForm
10   Const constSource As String = mconModuleName & ".DetachAll"
20   On Error GoTo DetachAll_Err
     'TraceCtl  "DetachAll"
30   oSub.Terminate
40   For Each df In oDockedForms
50      df.DetachFormWnd
60      Next
70   DetachAll_Exit:
80   On Error GoTo 0
90   Exit Sub
100   DetachAll_Err:
110   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", DetachAll", mconModuleName
120   Resume DetachAll_Exit
End Sub
'*****************************************************
Public Function Dock(ByRef df As DokNForm, Optional bHide As Boolean) As Boolean '4Matz:New
'*****************************************************
   ' Purpose   - Dock a form within host boundaries
   '** 12/04/06 Yorgi- The Splitter control now handles drawing requirements, here we just set window properties
   Dim lStyle              As Long
   Dim octl                As clsControl
   Dim ofrm                As Form
   Dim picHost             As PictureBox
10   On Error GoTo Dock_Err
20   Const constSource As String = mconModuleName & ".Dock"
     'TraceCtl  constSource & " begin"
30   If Not df Is Nothing Then
40      Set ofrm = df.DockedForm
50      If Not ofrm Is Nothing Then
           ' check if the form may dock here
60         If df.AllowDocking = False Then
              ' if not just show the form and good bye
70            ShowWindow ofrm.hWnd, SW_SHOWNORMAL
80         Else
90            If df.State <> DS_Docked Then
100               ofrm.Visible = False
                  'Add Host Container to the ContainedCntrls
110               Set picHost = df.HostContainer
                  ''''                  'keep a weak reference to DockedForm object in host picture container
                  ''''130               SetPropA picHost.hWnd, gconPROPERTY_DFPTR, ObjPtr(df)
120               SetParent ofrm.hWnd, picHost.hWnd 'Set our parent to the HostContainer
                  ' set the form's window style
130               lStyle = GetWindowLong(ofrm.hWnd, GWL_STYLE)
140               lStyle = lStyle Or WS_DLGFRAME Or WS_SYSMENU Or WS_OVERLAPPED Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS
150               lStyle = lStyle And Not (WS_MAXIMIZE Or WS_MINIMIZE Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME Or WS_CAPTION)
160               SetWindowLong ofrm.hWnd, GWL_STYLE, lStyle
                  ' set the form's window extended style
170               lStyle = GetWindowLong(ofrm.hWnd, GWL_EXSTYLE)
180               lStyle = lStyle And Not WS_EX_APPWINDOW
190               SetWindowLong ofrm.hWnd, GWL_EXSTYLE, lStyle
                  ' show docked window
200               df.State = DS_Docked            'change the dock form state
                  '-----------------------------------------
                  '-- Create a new virtual control
                  '-----------------------------------------
210               If GetParent(picHost.hWnd) <> UserControl.hWnd Then
220                  HostCtrlAdd octl, picHost, df.AttachToCtrlPtr, df.Align
230                  octl.DFKey = df.Key          'backward looking key from clsControl to DockedForm
                     '-----------------------------------------
                     '-- Modify ctlTitleBar and reload persistant values from DokNForm object
                     '-----------------------------------------
240                  df.VCtlKey = octl.Key        'back reference the virtual ctrl id
250                  octl.MinHeight = df.MinHeight
260                  octl.MinWidth = df.MinWidth
270                  With ctbTitlebar(octl.TbarIdx)
280                     If LenB(.Caption) = 0 Then
290                        .Caption = ofrm.Caption 'grab the form's caption
300                        End If
310                     If (df.Style And DSFloat) = DSFloat Then
320                        .TBarType = TBT_DEFAULT 'docking forms will have double stripe
330                     Else
340                        .TBarType = (df.TBarType Or TBT_SINGLESTRIPE) And Not TBT_DOUBLESTRIPE
350                        End If
360                     .Orientation = df.TBarPos ' and Position for each dockedform
370                     .CloseVisible = df.HasCloseButton
380                     End With
390                  End If
400               ShowControl df.VCtlKey, True    'open/show the valid docked control
410               Dock = True
420               End If
430            ofrm.Visible = True
440            End If
450         End If
460      End If
      'TraceCtl  constSource & " end"
470   Dock_Exit:
480   On Error GoTo 0
490   Exit Function
500   Dock_Err:
510   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", Dock", mconModuleName
520   Resume Dock_Exit
End Function
'*****************************************************
Public Function DockedForm(ByRef ofrm As Object) As DokNForm '4Matz:New
'*****************************************************
   ' Purpose   - Retrieves the docked form object
10   If Not ofrm Is Nothing Then
20      Set DockedForm = oDockedForms.ItemByHandle(ofrm.hWnd)
30      End If
End Function
'*****************************************************
Public Property Get Enable() As Boolean
Attribute Enable.VB_Description = "Returns/sets a value that determines whether all splitters are movable"
'*****************************************************
   ' Purpose    - Returns a value that determines whether all splitters are movable
10   Enable = mSplitters.Enable
End Property
'*****************************************************
Public Property Let Enable(blnEnable As Boolean)
'*****************************************************
   ' Purpose    - Sets a value that determines whether all splitters are movable
   ' Input      - blnEnable (the new Enable property value)
10   mSplitters.Enable = blnEnable
20   PropertyChanged mconEnable
30   If Ambient.UserMode = False Then Refresh     'only during design time
End Property
'*****************************************************
Public Property Get FillContainer() As Boolean
Attribute FillContainer.VB_Description = "Returns/sets a value that determines whether the ActiveX Control (along with all controls inside it) will automatically adjust its size to fill-up its container with respect to the margin properties"
'*****************************************************
   ' Purpose    - Returns a value that determines whether the ActiveX Control
   '              (along with all controls inside it) will automatically adjust its
   '              size to fill-up its container with respect to the margin
   '              properties
10   FillContainer = mblnFillContainer
End Property
'*****************************************************
Public Property Let FillContainer(blnFillContainer As Boolean)  '4Matz:Changed
'*****************************************************
   ' Purpose    - Sets a value that determines whether the ActiveX Control (along
   '              with all controls inside it) will automatically adjust its size
   '              to fill-up its container with respect to the margin properties
   ' Input      - blnFillContainer (the new FillContainer property value)
10   mblnFillContainer = blnFillContainer
20   PropertyChanged mconFillContainer
     ''''30   If mblnFillContainer Then Activate ' only is mblnFillContainer = true
End Property
'*****************************************************
Public Function FormAdd(ByRef ofrm As Object, Optional df As DokNForm, Optional oAttachToCtrl As Object, Optional Align As eDAlignProperty = DAlignLeft, Optional sKey As String, Optional Style As eDockStyles, Optional iPos As TBarOrientation, Optional bHasCloseButton As Boolean = True) As Boolean '4Matz:New
'*****************************************************
   ' Purpose   - Creates a docked form object and it's Host Container.  A new DoknForm objects is created
   '             only if it does not exist, otherwise we ShowControl the existing object (positions not changed)
   '             If you want to force a particular position, first make sure object is not .Closed,
   '             then use the MoveControl function for specific placement.
   Dim picHost             As PictureBox
10   Const constSource As String = mconModuleName & ".FormAdd"
20   On Error GoTo FormAdd_Err
30   Debug.Assert (Not ofrm Is Nothing)           'duh?
40   Debug.Assert (ofrm.WindowState = 0)          'sizing routines require sizeable form (duh*2)
50   Set df = oDockedForms.ItemByHandle(ofrm.hWnd)
60   If Not df Is Nothing Then
70      If LenB(df.VCtlKey) Then
80         ShowControl df.VCtlKey, True           'open/show the valid docked control
90         RaiseEvent FormAddComplete(ofrm.Name, df.Key, df.VCtlKey)
100         FormAdd = True
110      Else
            ' if not just show the form and good bye
120         ShowWindow ofrm.hWnd, SW_SHOWNORMAL
130         End If
140   Else                                        'dockedform not found, so build it
         ' if the form style was not furnished then set all styles available to the form
150      If Style <= 0 Then
160         Style = DSFloat Or DSLeft Or DSRight Or DSTop Or DSBottom
170         End If
         'check for a key name or default to form name
180      If LenB(sKey) = 0 Then sKey = ofrm.Name
190      With Parent.Controls
200         On Error Resume Next                  'Is host already created?
210         Set picHost = .Item(mconHostCtlPrefixName & ofrm.Name)
220         On Error GoTo FormAdd_Err
230         If picHost Is Nothing Then            'if not exist create it
               'Create a Picturebox to act as the form's Host Container while docked
240            Set picHost = .Add("vb.picturebox", mconHostCtlPrefixName & ofrm.Name)
               'set the defaults for all hosts
250            With picHost
260               .Enabled = True
270               .BorderStyle = 0
280               .TabStop = False
290               .Visible = True
300               End With
310            End If
320         End With
         ' add the form to the list of dockable forms
330      Set df = oDockedForms.Add(ofrm, picHost, mlngAppHosthWnd, Style, sKey, iPos, bHasCloseButton)
340      df.Align = Align
350      df.AttachToCtrlPtr = oAttachToCtrl
360      Set df.oIDF = Me                         'send interface messages back to me
         'Setup form attributes as a child docked object within its Picturebox Host Container, need to do before
         'HostCtrlAdd so correct minimum sizes are set
370      Dock df
380      RaiseEvent FormAddComplete(ofrm.Name, sKey, df.VCtlKey)
390      FormAdd = True
400      End If
410   FormAdd_Exit:
420   On Error GoTo 0
430   Exit Function
440   FormAdd_Err:
450   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", FormAdd", mconModuleName
460   Resume FormAdd_Exit
End Function
'*****************************************************
Public Function FormRemove(ByRef ofrm As Object, Optional bUnload As Boolean = True) As Boolean '4Matz:New
'*****************************************************
   ' Purpose   - Unloads the docked form object and it's Host Container
   Dim df                  As DokNForm
   Dim sName               As String
10   Const constSource As String = "FormRemove"
20   On Error GoTo FormRemove_Err
     'TraceCtl  constSource & " begin"
30   If Not ofrm Is Nothing Then                  'make sure the form is still valid
40      sName = ofrm.Name
        'TraceCtl  constSource & " Form:" & sName
50      Set df = oDockedForms.ItemByHandle(ofrm.hWnd)
60      If Not df Is Nothing Then
           'TraceCtl  constSource & " UnDock:" & sName
70         df.DetachFormWnd                       'stop subclassing
80         UnDock df, True
90         Set df = Nothing                       'try to de-reference so delete can occur
100         oDockedForms.Remove sName             'remove from collection
110         End If
120      If bUnload Then
            'TraceCtl  constSource & " Unload:" & sName
130         ofrm.Visible = False
140         Unload ofrm
150         Set ofrm = Nothing
160         End If
170      FormRemove = True
180      RaiseEvent FormRemoveComplete(sName)
190      End If
200   FormRemove_Exit:
210   On Error GoTo 0
      'TraceCtl  constSource & " end"
220   Exit Function
230   FormRemove_Err:
240   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", FormRemove", mconModuleName
250   Resume FormRemove_Exit
End Function
'*****************************************************
Private Function GetDropRect(sIdCtlSource As String, sIdCtlDestination As String, IdSplDestination As Long, udeTargetType As genmMoveDestination) As RECT   '4Matz:Changed
'*****************************************************
   ' Purpose    - Retrieves the drop guider rectangle
   ' Inputs     - * sIdCtlSource (the control's id that will be moved)
   '              * sIdCtlDestination (the control's id where the control sIdCtlSource will be moved to)
   '              * IdSplDestination (the splitter's id where the control sIdCtlSource will be moved to)
   '              * udeTargetType (the target type (an edge or a splitter) of the drop rect)
   Dim lngHeight           As Long               'indicating the drop guider rectangle's height
   Dim lngSplSize          As Long
   Dim lngWidth            As Long               'indicating the drop guider rectangle's width
   Dim octl                As clsControl         'for enumerating all virtual controls
   Dim octlMinHeight       As clsControl         'to get the value of lngHeight
   Dim octlMinWidth        As clsControl         'to get the value of lngWidth
   Dim oId                 As clsId              'for enumerating all Id in Ids collection
   Dim urecGetDropRect     As RECT
10   Const conMaxCtlAreaTaken = 0.5               'the maximum percentage area of the
     'other controls allowed to be taken
20   Const conCtlAreaTaken = 0.25                 'the percentage area of the target
     '    control that will be taken in
     '              move control action
     '              in Controls collection
30   Set octlMinWidth = New clsControl
40   Set octlMinHeight = New clsControl
50   lngSplSize = 0
60   Select Case udeTargetType
        Case mdControlTop
           '-- The cursor is on the control's top edge
70         With mVirtualControls(sIdCtlDestination)
80            If .IdSplTop <> gconUninitializedLong Then lngSplSize = mSplitters(.IdSplTop).Height \ 2
90            If (.Width < mVirtualControls(sIdCtlSource).MinWidth) Or ((.Height * conCtlAreaTaken) - lngSplSize < mVirtualControls(sIdCtlSource).MinHeight) Then
100               urecGetDropRect.Left = gconUninitializedLong
110            Else
120               mdlAPI.SetRect urecGetDropRect, .Left, .Top, .Right, (.Top + (.Height * conCtlAreaTaken))
130               End If
140            End With
150      Case mdControlRight
            '-- The cursor is on the control's right edge
160         With mVirtualControls(sIdCtlDestination)
170            If .IdSplRight <> gconUninitializedLong Then lngSplSize = mSplitters(.IdSplRight).Width \ 2
180            If ((.Width * conCtlAreaTaken) - lngSplSize < mVirtualControls(sIdCtlSource).MinWidth) Or (.Height < mVirtualControls(sIdCtlSource).MinHeight) Then
190               urecGetDropRect.Left = gconUninitializedLong
200            Else
210               mdlAPI.SetRect urecGetDropRect, (.Right - (.Width * conCtlAreaTaken)), .Top, .Right, .Bottom
220               End If
230            End With
240      Case mdControlBottom
            '-- The cursor is on the control's bottom edge
250         With mVirtualControls(sIdCtlDestination)
260            If .IdSplBottom <> gconUninitializedLong Then lngSplSize = mSplitters(.IdSplBottom).Height \ 2
270            If (.Width < mVirtualControls(sIdCtlSource).MinWidth) Or ((.Height * conCtlAreaTaken) - lngSplSize < mVirtualControls(sIdCtlSource).MinHeight) Then
280               urecGetDropRect.Left = gconUninitializedLong
290            Else
300               mdlAPI.SetRect urecGetDropRect, .Left, (.Bottom - (.Height * conCtlAreaTaken)), .Right, .Bottom
310               End If
320            End With
330      Case mdControlLeft
            '-- The cursor is on the control's left edge
340         With mVirtualControls(sIdCtlDestination)
350            If .IdSplLeft <> gconUninitializedLong Then lngSplSize = mSplitters(.IdSplLeft).Width \ 2
360            If ((.Width * conCtlAreaTaken) - lngSplSize < mVirtualControls(sIdCtlSource).MinWidth) Or (.Height < mVirtualControls(sIdCtlSource).MinHeight) Then
370               urecGetDropRect.Left = gconUninitializedLong
380            Else
390               mdlAPI.SetRect urecGetDropRect, .Left, .Top, (.Left + (.Width * conCtlAreaTaken)), .Bottom
400               End If
410            End With
420      Case mdEdgeTop
            '-- The cursor is on the top edge
            ' Get the height of the dropped control
430         octlMinHeight.Height = gconLngInfinite
440         For Each octl In mVirtualControls
450            If (Not octl.Closed) Then
460               If (octl <> mVirtualControls(sIdCtlSource)) Then
470                  If (octl.Bottom < octlMinHeight.Bottom) Then
480                     Set octlMinHeight = octl
490                     End If
500                  End If
510               End If
520            Next
530         lngHeight = GetMin(mVirtualControls(sIdCtlSource).Height, (octlMinHeight.Height * conMaxCtlAreaTaken) - mSplitters.Size)
            ' If the height is less than the minimum height, don't draw the drop guider
            '   rectangle, else draw the drop guider rectangle
540         If (lngHeight < octlMinHeight.MinHeight) Or (lngHeight < mVirtualControls(sIdCtlSource).MinHeight) Then
550            urecGetDropRect.Left = gconUninitializedLong
560         Else
570            mdlAPI.SetRect urecGetDropRect, 0, 0, mtypUCInside.Width, lngHeight
580            End If
590      Case mdEdgeRight
            '-- The cursor is on the right edge
            ' Get the width of the dropped control
600         octlMinWidth.Width = 0
610         For Each octl In mVirtualControls
620            If (Not octl.Closed) Then
630               If (octl <> mVirtualControls(sIdCtlSource)) Then
640                  If (octl.Left > octlMinWidth.Left) Then
650                     Set octlMinWidth = octl
660                     End If
670                  End If
680               End If
690            Next
700         lngWidth = GetMin(mVirtualControls(sIdCtlSource).Width, (octlMinWidth.Width * conMaxCtlAreaTaken) - mSplitters.Size)
            ' If the width is less than the minimum width, don't draw the drop guider
            '   rectangle, else draw the drop guider rectangle
710         If (lngWidth < octlMinWidth.MinWidth) Or (lngWidth < mVirtualControls(sIdCtlSource).MinWidth) Then
720            urecGetDropRect.Left = gconUninitializedLong
730         Else
740            mdlAPI.SetRect urecGetDropRect, mtypUCInside.Width - lngWidth, 0, mtypUCInside.Width, mtypUCInside.Height
750            End If
760      Case mdEdgeBottom
            '-- The cursor is on the bottom edge
            ' Get the height of the dropped control
770         octlMinHeight.Height = 0
780         For Each octl In mVirtualControls
790            If (Not octl.Closed) Then
800               If (octl <> mVirtualControls(sIdCtlSource)) Then
810                  If (octl.Top > octlMinHeight.Top) Then
820                     Set octlMinHeight = octl
830                     End If
840                  End If
850               End If
860            Next
870         lngHeight = GetMin(mVirtualControls(sIdCtlSource).Height, (octlMinHeight.Height * conMaxCtlAreaTaken) - mSplitters.Size)
            ' If the height is less than the minimum height, don't draw the drop guider
            '   rectangle, else draw the drop guider rectangle
880         If (lngHeight < octlMinHeight.MinHeight) Or (lngHeight < mVirtualControls(sIdCtlSource).MinHeight) Then
890            urecGetDropRect.Left = gconUninitializedLong
900         Else
910            mdlAPI.SetRect urecGetDropRect, 0, mtypUCInside.Height - lngHeight, mtypUCInside.Width, mtypUCInside.Height
920            End If
930      Case mdEdgeLeft
            '-- The cursor is on the left edge
            ' Get the width of the dropped control
940         octlMinWidth.Width = gconLngInfinite
950         For Each octl In mVirtualControls
960            If (Not octl.Closed) And (octl <> mVirtualControls(sIdCtlSource)) And (octl.Right < octlMinWidth.Right) Then Set octlMinWidth = octl
970            Next
980         lngWidth = GetMin((octlMinWidth.Width * conMaxCtlAreaTaken) - mSplitters.Size, mVirtualControls(sIdCtlSource).Width)
            ' If the width is less than the minimum width, don't draw the drop guider
            '   rectangle, else draw the drop guider rectangle
990         If (lngWidth < octlMinWidth.MinWidth) Or (lngWidth < mVirtualControls(sIdCtlSource).MinWidth) Then
1000            urecGetDropRect.Left = gconUninitializedLong
1010         Else
1020            mdlAPI.SetRect urecGetDropRect, 0, 0, lngWidth, mtypUCInside.Height
1030            End If
1040      Case mdSplitter
             '-- The cursor is on a splitter
1050         Select Case mSplitters(IdSplDestination).Orientation
                Case orHorizontal
                   ' Get the height of the dropped control
1060               octlMinHeight.Height = gconLngInfinite
1070               For Each oId In mSplitters(IdSplDestination).IdsCtlTop
1080                  If (mVirtualControls.ItemNo(oId.Id) <> mVirtualControls(sIdCtlSource)) Then
1090                     If (mVirtualControls.ItemNo(oId.Id).Height < octlMinHeight.Height) Then
1100                        Set octlMinHeight = mVirtualControls.ItemNo(oId.Id)
1110                        End If
1120                     End If
1130                  Next
1140               For Each oId In mSplitters(IdSplDestination).IdsCtlBottom
1150                  If (mVirtualControls.ItemNo(oId.Id) <> mVirtualControls(sIdCtlSource)) Then
1160                     If (mVirtualControls.ItemNo(oId.Id).Height < octlMinHeight.Height) Then
1170                        Set octlMinHeight = mVirtualControls.ItemNo(oId.Id)
1180                        End If
1190                     End If
1200                  Next
1210               lngHeight = GetMin(mVirtualControls(sIdCtlSource).Height, (octlMinHeight.Height * conMaxCtlAreaTaken) - mSplitters.Size)
                   ' If the height is less than the minimum height, don't draw the drop
                   '   guider rectangle, else draw the drop guider rectangle
1220               If (lngHeight < octlMinHeight.MinHeight) Or (lngHeight < mVirtualControls(sIdCtlSource).MinHeight) Then
1230                  urecGetDropRect.Left = gconUninitializedLong
1240               Else
1250                  With mSplitters(IdSplDestination)
1260                     mdlAPI.SetRect urecGetDropRect, .Left, .Top - lngHeight, .Right, .Bottom + lngHeight
1270                     End With
1280                  End If
1290            Case orVertical
                   ' Get the width of the dropped control
1300               octlMinWidth.Width = gconLngInfinite
1310               For Each oId In mSplitters(IdSplDestination).IdsCtlLeft
1320                  If (mVirtualControls.ItemNo(oId.Id) <> mVirtualControls(sIdCtlSource)) Then
1330                     If (mVirtualControls.ItemNo(oId.Id).Width < octlMinWidth.Width) Then
1340                        Set octlMinWidth = mVirtualControls.ItemNo(oId.Id)
1350                        End If
1360                     End If
1370                  Next
1380               For Each oId In mSplitters(IdSplDestination).IdsCtlRight
1390                  If (mVirtualControls.ItemNo(oId.Id) <> mVirtualControls(sIdCtlSource)) Then
1400                     If (mVirtualControls.ItemNo(oId.Id).Width < octlMinWidth.Width) Then
1410                        Set octlMinWidth = mVirtualControls.ItemNo(oId.Id)
1420                        End If
1430                     End If
1440                  Next
1450               lngWidth = GetMin(mVirtualControls(sIdCtlSource).Width, (octlMinWidth.Width * conMaxCtlAreaTaken) - mSplitters.Size)
                   ' If the width is less than the minimum width, don't draw the drop
                   '   guider rectangle, else draw the drop guider rectangle
1460               If (lngWidth < octlMinWidth.MinWidth) Or (lngWidth < mVirtualControls(sIdCtlSource).MinWidth) Then
1470                  urecGetDropRect.Left = gconUninitializedLong
1480               Else
1490                  With mSplitters(IdSplDestination)
1500                     mdlAPI.SetRect urecGetDropRect, .Left - lngWidth, .Top, .Right + lngWidth, .Bottom
1510                     End With
1520                  End If
1530            End Select
1540      End Select
1550   Set octlMinWidth = Nothing
1560   Set octlMinHeight = Nothing
1570   GetDropRect = urecGetDropRect
End Function
'*****************************************************
Private Sub GetDropTarget(ByRef blnTargetValid As Boolean, ByRef udeTargetType As genmMoveDestination, ByRef sIdCtl As String, ByRef lngIdSpl As Long, ByRef uposCursor As POINTAPI) '4Matz:Changed
'*****************************************************
   ' Purpose    - Retrieves the drop guider target type based on the current mouse position
   ' Returns    - * blnTargetValid (indicating whether the current mouse position is on a valid target)
   '              * udtTargetType (the target type- an edge or a control's edge or a splitter of the drop rect)
   '              * sIdCtl (the target control's id)
   '              * lngIdSpl (the target splitter's id)
   '** 01/25/07 Yorgi- Return valid target only if cursor is inside the Usercontrol
   '** 01/26/07 Yorgi- Performance & cleanup
   Dim lBottom             As Long
   Dim lCursorX            As Long
   Dim lCursorY            As Long
   Dim lLeft               As Long
   Dim lRight              As Long
   Dim lSplitrSize         As Long               'relative to the DoknSplitz control
   Dim lTop                As Long
   Dim octl                As clsControl         'for enumerating all virtual controls
   Dim ospl                As clsSplitter        'for enumerating all virtual splitters
10   Const conCtlAreaTaken = 0.1                  'the percentage area of the target control that will be taken in move control action
20   blnTargetValid = False
30   udeTargetType = gconUninitializedLong        'means no valid target found
40   sIdCtl = vbNullString
50   uposCursor = GetCursorRelPos(UserControl.hWnd, lCursorX, lCursorY) 'get the current cursor position
     'TraceCtl  "ctbTitleBar_MoveEnd X:" & uposCursor.X & ", Y:" & uposCursor.Y & ", PtInRect:" & PtInRect(mrctUserControl, lCursorX, lCursorY)
     '** Yorgi- Return valid target only if cursor is inside the Usercontrol
60   If PtInRect(mrctUserControl, lCursorX, lCursorY) <> 0 Then
70      lCursorX = uposCursor.X                   'localize the X value
80      lCursorY = uposCursor.Y                   'localize the Y value
90      lngIdSpl = gconUninitializedLong
        '-- Check if cursor is outside DoknSplitz control rect
100      lSplitrSize = mSplitters.Size
110      If lCursorX <= lSplitrSize Then          'Is Left edge
120         udeTargetType = mdEdgeLeft
130      ElseIf lCursorX >= mtypUCInside.Width - lSplitrSize Then 'Is Right edge
140         udeTargetType = mdEdgeRight
150      ElseIf lCursorY <= lSplitrSize Then      'Is Top edge
160         udeTargetType = mdEdgeTop
170      ElseIf lCursorY >= mtypUCInside.Height - lSplitrSize Then 'Is Bottom edge
180         udeTargetType = mdEdgeBottom
190      Else                                     '-- Check whether the cursor is on the edge of a control
200         For Each octl In mVirtualControls
210            If Not octl.Closed Then
                  '-- YorgiPerf: VB compound if statements are not as efficient as C++ so break them apart
220               If (octl.Left <= lCursorX) Then
230                  If (lCursorX <= octl.Right) Then
240                     If (octl.Top <= lCursorY) Then
250                        If (lCursorY <= octl.Bottom) Then
260                           If ((lCursorX <= octl.Left + (octl.Width * conCtlAreaTaken)) Or (lCursorX >= octl.Right - (octl.Width * conCtlAreaTaken)) Or (lCursorY <= octl.Top + (octl.Height * conCtlAreaTaken)) Or (lCursorY >= octl.Bottom - (octl.Height * conCtlAreaTaken))) Then
270                              sIdCtl = octl.Key
                                 '-- YorgiPerf: do the math one time!!!
280                              lTop = lCursorY - octl.Top
290                              lRight = octl.Right - lCursorX
300                              lBottom = octl.Bottom - lCursorY
310                              lLeft = lCursorX - octl.Left
320                              Select Case GetMin(lTop, lRight, lBottom, lLeft)
                                    Case lTop
330                                    udeTargetType = mdControlTop
340                                 Case lRight
350                                    udeTargetType = mdControlRight
360                                 Case lBottom
370                                    udeTargetType = mdControlBottom
380                                 Case lLeft
390                                    udeTargetType = mdControlLeft
400                                 End Select
410                              Exit For
420                              End If
430                           End If
440                        End If
450                     End If
460                  End If
470               End If
480            Next
490         End If
         '-- Check whether the cursor is on a splitter
500      If udeTargetType = gconUninitializedLong Then
510         For Each ospl In mSplitters
               '-- YorgiPerf: VB compound if statements are not as efficient as C++ so break them apart
520            If (ospl.Left <= lCursorX) Then
530               If (lCursorX <= ospl.Right) Then
540                  If (ospl.Top <= lCursorY) Then
550                     If (lCursorY <= ospl.Bottom) Then
560                        lngIdSpl = ospl.Id
570                        udeTargetType = mdSplitter
580                        Exit For
590                        End If
600                     End If
610                  End If
620               End If
630            Next
640         End If
650      End If
660   blnTargetValid = Not (udeTargetType = gconUninitializedLong)
End Sub
'*****************************************************
Private Function GetParentHwnd() As Long '4Matz:New
'*****************************************************
   ' Purpose    - Returns Parent or MDIChild hWnd
   Dim lhWnd               As Long
   Dim lHwnd2              As Long
10   lhWnd = Extender.Parent.hWnd                 'get the Hwnd of our parent
20   lHwnd2 = FindWindowEx(lhWnd, 0, "MDIClient", ByVal 0&) 'see if we are in an MDI form
30   If lHwnd2 Then                               'use the MDIClient window
40      GetParentHwnd = lHwnd2
50   Else
60      GetParentHwnd = lhWnd                     ' use the parent window
70      End If
End Function
'*****************************************************
Private Function GetUCInnerDimensions(ByRef Innerds As typUCInnerDimensions) '4Matz:New
'*****************************************************
   ' Purpose    - Returns the Usercontrol's inside dimensions with respect to the space used by the Slider
   Dim lAlign              As Long
10   Innerds.Left = 0
20   Innerds.Top = 0
30   Innerds.Width = UserControl.ScaleWidth
40   Innerds.Height = UserControl.ScaleHeight
     'TraceCtl  "..GetUCInnerDimensions Left:" & Innerds.Left & " Top:" & Innerds.Top & " Width:" & Innerds.Width & " Height:" & Innerds.Height
     ''''50        If mlngSliderThickness > 0 Then    ' check for aligned control
     ''''60           lAlign = Extender.Align
     ''''70           Select Case lAlign
     ''''                Case vbAlignLeft, vbAlignRight
     ''''80                 Innerds.Width = Innerds.Width - mlngSliderThickness
     ''''90                 If lAlign = vbAlignRight Then Innerds.Left = mlngSliderThickness
     ''''100              Case vbAlignTop, vbAlignBottom
     ''''110                 Innerds.Height = Innerds.Height - mlngSliderThickness
     ''''120              End Select
     ''''130           End If
End Function
'*****************************************************
Private Function HasStretched(ByRef sngXScale As Single, ByRef sngYScale As Single) As Boolean
Attribute HasStretched.VB_Description = "Returns a valid x- and y- coordinate scale"
'*****************************************************
   ' Purpose    - Returns a valid x- and y- coordinate scale
   Dim octl                As clsControl         'for enumerating all virtual controls
10   If mVirtualControls.Count Then
20      sngXScale = mtypUCInside.Width / mVirtualControls.Width
30      sngYScale = mtypUCInside.Height / mVirtualControls.Height
40      For Each octl In mVirtualControls
50         If Not octl.Closed Then
60            If octl.Width * sngXScale < octl.MinWidth Then sngXScale = 1
70            If octl.Height * sngYScale < octl.MinHeight Then sngYScale = 1
80            End If
90         Next
100      HasStretched = (Abs(sngXScale - 1) > 0.001) Or (Abs(sngYScale - 1) > 0.001)
110      End If
      'TraceCtl  "HasStretched Width:" & sngXScale & " Height:" & sngYScale
End Function
'*****************************************************
Private Function HostCtrlAdd(ByRef oVirtCtl As clsControl, ByRef oRTCtrl As PictureBox, ByRef oAttachToCtrl As Control, lAttachWhere As eDAlignProperty) As Long '4Matz:New
'*****************************************************
   ' Purpose    - Add runtime Host control to the controls collection and rebuild all splitters
   ' Effect     - * If successed, as specified
   '              * Otherwise, the custom error message has been raised
   ' 11/09/06 Yorgi - Create function
   ' 11/24/06 Yorgi : Allow add by position
   'Change your RunTimeCtrl's parent to the splitter. This causes an element to be added to the
   ' ContainedCntrls collection.
   Dim lHeight             As Long
   Dim lLeft               As Long
   Dim lTop                As Long
   Dim lWidth              As Long
10   On Error GoTo HostCtrlAdd_Err
20   SetParent oRTCtrl.hWnd, UserControl.hWnd
     'TraceCtl  "HostCtrlAdd " & oRTCtrl.Name & " hWnd(" & oRTCtrl.hWnd & ") FormParentHwnd(" & UserControl.hWnd & ")"
     'manipulate dimensions so desired positions and splitters are created
30   Select Case lAttachWhere
        Case DAlignLeft, DAlignRight              'Set Left & Right aligned controls
40         If Not oAttachToCtrl Is Nothing Then
50            lTop = oAttachToCtrl.Top
60            lHeight = oAttachToCtrl.Height
70            oAttachToCtrl.Width = oAttachToCtrl.Width \ 2
80            lWidth = oAttachToCtrl.Width
90            If lAttachWhere = DAlignLeft Then
100               lLeft = oAttachToCtrl.Left
110               oAttachToCtrl.Left = lLeft + lWidth
120            Else
130               lLeft = oAttachToCtrl.Left + oAttachToCtrl.Width
140               End If
150         Else
               'no adjacent control selected so we have to use the DockNSplit actual size
160            lHeight = mtypUCInside.Height
170            lWidth = mtypUCInside.Width * 0.25
180            If lAttachWhere = DAlignRight Then
190               lLeft = mtypUCInside.Width - lWidth
200               End If
210            End If
220      Case Else                                'Set Top & Bottom aligned controls
230         If Not oAttachToCtrl Is Nothing Then
240            lWidth = oAttachToCtrl.Width
250            lLeft = oAttachToCtrl.Left
260            lHeight = oAttachToCtrl.Height \ 2
270            oAttachToCtrl.Height = lHeight
280            If lAttachWhere = DAlignTop Then
290               lTop = oAttachToCtrl.Top
300               oAttachToCtrl.Top = lTop + lHeight
310            Else
320               lTop = oAttachToCtrl.Top + oAttachToCtrl.Height
330               End If
340         Else
               'no adjacent control selected so we have to use the DockNSplit actual size
350            lWidth = mtypUCInside.Width
360            lHeight = mtypUCInside.Height * 0.25
370            If lAttachWhere = DAlignBottom Then
380               lTop = mtypUCInside.Height - lHeight
390               End If
400            End If
410      End Select
420   oRTCtrl.Move lLeft, lTop, lWidth, lHeight   'move host to the display location
430   HostCtrlAdd = VCtrlAdd(oRTCtrl, oVirtCtl)   'add host control to the VirtCtls collection
440   HostCtrlAdd_Exit:
450   On Error GoTo 0
460   Exit Function
470   HostCtrlAdd_Err:
480   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", HostCtrlAdd", mconModuleName
490   Resume HostCtrlAdd_Exit
End Function
'*****************************************************
Private Sub HostCtrlRemove(ByRef oRTCtrl As Control) '4Matz:New
'*****************************************************
   ' Purpose    - Remove runtime control/form from the controls collection
   ' Effect     - * If successed, as specified
   '              * Otherwise, the custom error message has been raised
   ' 12/05/06 Yorgi - Create function
   'Change your RunTimeCtrl's parent to the parent form.  This should also remove object
   'from the ContainedCntrls collection
10   On Error GoTo HostCtrlRemove_Err
     'TraceCtl  "HostCtrlRemove Host:" & oRTCtrl.Name
20   If GetParent(oRTCtrl.hWnd) <> Parent.hWnd Then
30      SetParent oRTCtrl.hWnd, Parent.hWnd       'switch parent hWnd ptr from Usercontrol back to Parent.Control
40      End If
50   oRTCtrl.Visible = False                      'hide the control
60   Parent.Controls.Remove oRTCtrl.Name          'remove from the Parents control collection
70   HostCtrlRemove_Exit:
80   On Error GoTo 0
90   Exit Sub
100   HostCtrlRemove_Err:
110   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", HostCtrlRemove", mconModuleName
120   Resume HostCtrlRemove_Exit
End Sub
'*****************************************************
Private Sub IDoknForm_WinEvent(hWnd As Long, uMsg As Long, df As DokNForm, wParam As Long, lParam As Long) '4Matz:New
'*****************************************************
   ' Purpose    - An interface postback method for special handling of form events
   Dim bResult             As Boolean
   Dim lId                 As Integer
   Dim rcMove              As RECT
10   On Error GoTo IDoknForm_WinEvent_Err
     'TraceCtl  "IDoknForm_WinEvent hWnd:" & CStr(hWnd)
20   lId = mVirtualControls(df.VCtlKey).TbarIdx
30   Select Case uMsg
        Case WM_MOVING
           'while undocked, check for a new parking spot.  If docked, the moves are handled by ctbTitleBar events
40         ctbTitleBar_Move lId, 0, True          'hitchhiking on the ctbTitleBar events for moving docked controls
50      Case WM_ENTERSIZEMOVE
60         df.DockedForm.Hide                     'now we can better see the windows drag window
70         ctbTitleBar_MoveBegin lId, 0, True     'hitchhiking on the ctbTitleBar events for moving docked controls
80      Case WM_EXITSIZEMOVE
90         ctbTitleBar_MoveEnd lId, 0, df, bResult 'hitchhiking on the ctbTitleBar events for moving docked controls
100         If Not bResult Then
110            df.DockedForm.Show                 'apparently no parking spot available, so unHide the form
120            End If
130      Case WM_DESTROY
            'should only happen if user closes an undocked form, or programmer unloads the form
140         bResult = (df.State = DS_Docked)
150         HostCtrlRemove df.HostContainer       'remove Host Control from ContainedControls
160         VCtrlRemove df.VCtlKey                'remove host control from the VirtCtls collection
170         oDockedForms.Remove df.Key            'remove from DockedForms collection
180         If bResult Then VCtrlManager          'if we were docked, have to refresh DoknSplitz display
190      End Select
200   IDoknForm_WinEvent_Exit:
210   On Error GoTo 0
220   Exit Sub
230   IDoknForm_WinEvent_Err:
240   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", IDoknForm_WinEvent", mconModuleName
250   Resume IDoknForm_WinEvent_Exit
End Sub
'*****************************************************
Private Sub InitSlider(eAlign As AlignConstants) '4Matz:New
'*****************************************************
   ' Purpose    - Changes in Align property determine Slider's visibility
10   If eAlign = vbAlignNone Then
20      picSlider.Visible = False
30      mlngSliderThickness = 0
40   Else
50      If eAlign = vbAlignLeft Or eAlign = vbAlignRight Then
60         oSlider.Orientation = cSPLTOrientationVertical
70         picSlider.MousePointer = vbSizeWE
80      Else
90         oSlider.Orientation = cSPLTOrientationHorizontal
100         picSlider.MousePointer = vbSizeNS
110         End If
120      mlngSliderThickness = mSplitters.Size
130      picSlider.Visible = True
140      End If
      'TraceCtl  "InitSlider eAlign:" & eAlign
End Sub
'*****************************************************
Private Function IsRectNearSource(sIdCtlSource As String, sIdCtlDestination As String, IdSplDestination As Long, udeTargetType As genmMoveDestination) As Boolean
'*****************************************************
   ' Purpose    - Returns a value indicating whether the target move action
   '              udeTargetType is near the control IdCtlSource which will be moved
   ' Inputs     - * IdCtlSource (the id of the control which will be moved)
   '              * IdCtlDestination (the id of the control where the control
   '                                  IdCtlSource will be moved to [for MoveTo =
   '                                  mdControlTop, mdControlRight, mdControlBottom
   '                                  or mdControlLeft)
   '              * IdSplDestination (the id of the splitter where the control
   '                                  IdControlSource will be moved to [for
   '                                  MoveTo = mdSplitter])
   '              * udeTargetType (the type of the area [an edge, a control's edge
   '                               or a splitter] where the control IdControl will
   '                               be moved to)
   Dim blnIsRectNearSource As Boolean
10   Select Case udeTargetType
        Case mdControlTop, mdControlRight, mdControlBottom, mdControlLeft
20         blnIsRectNearSource = (sIdCtlSource = sIdCtlDestination)
30      Case mdSplitter
40         blnIsRectNearSource = False
50         With mVirtualControls(sIdCtlSource)
60            If .IdSplTop <> gconUninitializedLong Then blnIsRectNearSource = blnIsRectNearSource Or ((IdSplDestination = .IdSplTop) And (mSplitters(.IdSplTop).IdsCtlBottom.Count = 1))
70            If .IdSplRight <> gconUninitializedLong Then blnIsRectNearSource = blnIsRectNearSource Or ((IdSplDestination = .IdSplRight) And (mSplitters(.IdSplRight).IdsCtlLeft.Count = 1))
80            If .IdSplBottom <> gconUninitializedLong Then blnIsRectNearSource = blnIsRectNearSource Or ((IdSplDestination = .IdSplBottom) And (mSplitters(.IdSplBottom).IdsCtlTop.Count = 1))
90            If .IdSplLeft <> gconUninitializedLong Then blnIsRectNearSource = blnIsRectNearSource Or ((IdSplDestination = .IdSplLeft) And (mSplitters(.IdSplLeft).IdsCtlRight.Count = 1))
100            End With
110      End Select
120   IsRectNearSource = blnIsRectNearSource
End Function
'*****************************************************
Private Function IsSolid(Optional blnIncludeSplitter As Boolean = True) As Boolean '4Matz:Changed
'*****************************************************
   ' Purpose    - Returns a value indicating whether the virtual controls and splitters are solid
   ' Input      - blnIncludeSplitter (indicating whether the splitters are included to determine the solid state)
   '-- YorgiPerf: VB compound if statements are not as efficient as C++ so break them apart
   Dim lngExtent           As Long               'total extent of the virtual controls and splitters
   Dim lngSplBottomHeight  As Long               'the height of the virtual splitter on the bottom-side of the current enumerated virtual control
   Dim lngSplLeftWidth     As Long               'the width of the virtual splitter on theleft-side of the current enumerated virtual control
   Dim lngSplRightWidth    As Long               'the width of the virtual splitter on the right-side of the current enumerated virtual control
   Dim lngSplTopHeight     As Long               'the height of the virtual splitter on the top-side of the current enumerated virtual control
   Dim octl                As clsControl         'for enumerating all virtual controls in Controls collection
10   lngExtent = 0
20   For Each octl In mVirtualControls
30      If Not octl.Closed Then
40         lngSplTopHeight = 0
50         lngSplRightWidth = 0
60         lngSplBottomHeight = 0
70         lngSplLeftWidth = 0
80         If blnIncludeSplitter Then             'only need to check one time!
90            If octl.IdSplTop <> gconUninitializedLong Then
100               lngSplTopHeight = mSplitters(octl.IdSplTop).Height
110               End If
120            If octl.IdSplRight <> gconUninitializedLong Then
130               lngSplRightWidth = mSplitters(octl.IdSplRight).Width
140               End If
150            If octl.IdSplBottom <> gconUninitializedLong Then
160               lngSplBottomHeight = mSplitters(octl.IdSplBottom).Height
170               End If
180            If octl.IdSplLeft <> gconUninitializedLong Then
190               lngSplLeftWidth = mSplitters(octl.IdSplLeft).Width
200               End If
210            End If
220         lngExtent = lngExtent + ((octl.Width + (lngSplLeftWidth \ 2) + (lngSplRightWidth \ 2)) * (octl.Height + (lngSplTopHeight \ 2) + (lngSplBottomHeight \ 2)))
230         End If
240      Next
250   IsSolid = (lngExtent = 0) Or (lngExtent = (mVirtualControls.Width * mVirtualControls.Height))
End Function
'*****************************************************
Private Sub iSubclass_WndProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As eMsg, ByVal wParam As Long, ByVal lParam As Long) '4Matz:New
'*****************************************************
10   On Error GoTo iSubclass_WndProc_Err
20   Select Case uMsg
        Case WM_ACTIVATE
           '-- This subclassing is used to handle the possibility of the user
           '   swithing to another application while dragging the splitter or while
           '   moving the control
30         If wParam = WA_INACTIVE Then
40            If mblnDragSplitter Then picSplitter_MouseUp mScIndex, mScButton, mScShift, mScX, mScY
50            If crecControl.Visible Then
60               ctbTitleBar_MoveEnd mScIndex, mScShift, Nothing, bBefore 'bBefore is not used anyway
70               End If
80            End If
90      Case WM_SIZE, WM_SHOWWINDOW
           '-- In VB Splitter, developers need to add one line of code in their form
           '   resize event to call the Activate method. Now with this subclassing,
           '   there is no need to add any code to form to use basic features of
           '   Control Manager ActiveX Control
           '** 02/22/07 Yorgi- Fix(iSubclass_WndProc):Don't call Activate on SIZE_MINIMIZE it's MADNESS
100         If Not (wParam = SIZE_MINIMIZED) Then Activate
110         If uMsg = WM_SHOWWINDOW Then
120            oSub.DelMsg mlngHwndParent, WM_SHOWWINDOW, MSG_AFTER
130            End If
140      End Select
150   iSubclass_WndProc_Exit:
160   On Error GoTo 0
170   Exit Sub
180   iSubclass_WndProc_Err:
190   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", iSubclass_WndProc", mconModuleName
200   Resume iSubclass_WndProc_Exit
End Sub
'*****************************************************
Public Property Let LiveUpdate(blnLiveUpdate As Boolean)
'*****************************************************
   ' Purpose    - Sets a value that determines whether the controls should be
   '              resized as a splitter is moved
   ' Input      - blnLiveUpdate (the new LiveUpdate property value)
10   mSplitters.LiveUpdate = blnLiveUpdate
20   PropertyChanged mconLiveUpdate
End Property
'*****************************************************
Public Property Get LiveUpdate() As Boolean
Attribute LiveUpdate.VB_Description = "Returns/sets a value that determines whether the controls should be resized as a splitter is moved"
'*****************************************************
   ' Purpose    - Returns a value that determines whether the controls should be
   '              resized as a splitter is moved
10   LiveUpdate = mSplitters.LiveUpdate
End Property
'*****************************************************
Public Property Get MarginBottom() As Long
Attribute MarginBottom.VB_Description = "Returns/sets the bottom margin of the ActiveX Control from its container"
'*****************************************************
   ' Purpose    - Returns the bottom margin of the ActiveX Control from its
   '              container
10   MarginBottom = mlngMarginBottom
End Property
'*****************************************************
Public Property Let MarginBottom(lngMarginBottom As Long)
'*****************************************************
   ' Purpose    - Sets the bottom margin of the ActiveX Control from its container
   ' Input      - lngMarginBottom (the new MarginBottom property value)
10   mlngMarginBottom = lngMarginBottom
20   PropertyChanged mconMarginBottom
30   Activate
End Property
'*****************************************************
Public Property Get MarginLeft() As Long
Attribute MarginLeft.VB_Description = "Returns/sets the left margin of the ActiveX Control from its container"
'*****************************************************
   ' Purpose    - Returns the left margin of the ActiveX Control from its container
10   MarginLeft = mlngMarginLeft
End Property
'*****************************************************
Public Property Let MarginLeft(lngMarginLeft As Long)
'*****************************************************
   ' Purpose    - Sets the left margin of the ActiveX Control from its container
   ' Input      - lngMarginLeft (the new MarginLeft property value)
10   mlngMarginLeft = lngMarginLeft
20   PropertyChanged mconMarginLeft
30   Activate
End Property
'*****************************************************
Public Property Let MarginRight(lngMarginRight As Long)
'*****************************************************
   ' Purpose    - Sets the right margin of the ActiveX Control from its container
   ' Input      - lngMarginRight (the new MarginRight property value)
10   mlngMarginRight = lngMarginRight
20   PropertyChanged mconMarginRight
30   Activate
End Property
'*****************************************************
Public Property Get MarginRight() As Long
Attribute MarginRight.VB_Description = "Returns/sets the right margin of the ActiveX Control from its container"
'*****************************************************
   ' Purpose    - Returns the right margin of the ActiveX Control from its
   '              container
10   MarginRight = mlngMarginRight
End Property
'*****************************************************
Public Property Get MarginTop() As Long
Attribute MarginTop.VB_Description = "Returns/sets the top margin of the ActiveX Control from its container"
'*****************************************************
   ' Purpose    - Returns the top margin of the ActiveX Control from its container
10   MarginTop = mlngMarginTop
End Property
'*****************************************************
Public Property Let MarginTop(lngMarginTop As Long)
'*****************************************************
   ' Purpose    - Sets the top margin of the ActiveX Control from its container
   ' Input      - lngMarginTop (the new MarginTop property value)
10   mlngMarginTop = lngMarginTop
20   PropertyChanged mconMarginTop
30   Activate
End Property
'*****************************************************
Public Function MoveControl(sIdControlSource As String, ByVal MoveTo As genmMoveDestination, Optional sIdControlDestination As String = vbNullString, Optional IdSplitterDestination As Long = gconUninitializedLong) As Boolean
Attribute MoveControl.VB_Description = "Moves a control to certain area"
'*****************************************************
   ' Purpose    - Moves a control to certain area
   ' Effects    - * If successful, the control has been moved
   '              * If control IdControl or splitter IdSplitter doesn't exist, a
   '                run-time error has been generated
   '              * otherwise, no effect
   ' Inputs     - * IdControlSource (A value that uniquely identifies the source
   '                                 control the developer want to move)
   '              * MoveTo (A value indicating the area type where the source
   '                        control will be moved to)
   '              * IdControlDestination (A value that uniquely identifies the
   '                                      destination control the developer want to
   '                                      move the source control to. This input is
   '                                      required if the area type indicated by
   '                                      MoveTo input is a control.)
   '              * IdSplitterDestination (A value that uniquely identifies the
   '                                       splitter the developer want to move the
   '                                       source control to. This input is
   '                                       required only if the are type indicated
   '                                       by MoveTo is a splitter)
   ' Return     - bSuccess (a returned value that determines whether the MoveControl
   '                       method is successful)
   '              in Controls collection size and position in case the
   'Control Manager couldn't be rebuilt
   Dim bSuccess            As Boolean
   Dim octl                As clsControl         'for enumerating all virtual controls
   Dim udeRemoveHeapDirection As genmRemoveHeapDirection
   Dim urecControlBackup() As RECT               'backup of the Controls collection's
   Dim urecDrop            As RECT               'indicating the drop guider rectangle size and and position
10   On Error GoTo MoveControl_Err
20   If Not mVirtualControls.IsExist(sIdControlSource) Then
30      bSuccess = False
40      SecureRaiseError errIdControl, "MoveControl"
50   ElseIf (MoveTo = mdSplitter) And (Not mSplitters.IsExist(IdSplitterDestination)) Then 'NOT NOT...
60      bSuccess = False
70      SecureRaiseError errIdSplitter, "MoveControl"
80   ElseIf mVirtualControls(sIdControlSource).Closed Then
90      bSuccess = False
100      SecureRaiseError errMoveControlClosed, "MoveControl"
110   Else
120      urecDrop = GetDropRect(sIdControlSource, sIdControlDestination, IdSplitterDestination, MoveTo)
130      If urecDrop.Left = gconUninitializedLong Then
140         bSuccess = False
150         If Not crecControl.Visible Then SecureRaiseError errMoveControlRoom, "MoveControl"
160         crecControl.Visible = False
170      Else
180         crecControl.Visible = False
            '-- Backup the controls position in case the Control Manager couldn't be rebuilt
190         mVirtualControls.Backup
            '-- Move the virtual control IdControl
200         With mVirtualControls(sIdControlSource)
210            .Left = urecDrop.Left
220            .Top = urecDrop.Top
230            .Right = urecDrop.Right
240            .Bottom = urecDrop.Bottom
250            End With
            '-- Re-arrange the other virtual controls
260         Select Case MoveTo
               Case mdControlTop, mdControlBottom, mdEdgeTop, mdEdgeBottom
270               udeRemoveHeapDirection = rhdVertical
280            Case mdControlLeft, mdControlRight, mdEdgeLeft, mdEdgeRight
290               udeRemoveHeapDirection = rhdHorizontal
300            Case mdSplitter
310               Select Case mSplitters(IdSplitterDestination).Orientation
                     Case orHorizontal
320                     udeRemoveHeapDirection = rhdVertical
330                  Case orVertical
340                     udeRemoveHeapDirection = rhdHorizontal
350                  End Select
360            End Select
370         mVirtualControls.RemoveHeap sIdControlSource, True, udeRemoveHeapDirection
380         mVirtualControls.Compact
390         mVirtualControls.RemoveHoles
400         bSuccess = IsSolid(False)
410         If bSuccess Then
               '-- Rebuild the splitters and applies the virtual controls and splitters
420            VCtrlManager False
430         Else
               '-- Restore the controls' position and size
440            mVirtualControls.Restore
450            bSuccess = Refresh
460            End If
470         End If
480      End If
490   MoveControl = bSuccess
500   MoveControl_Exit:
510   On Error GoTo 0
520   Exit Function
530   MoveControl_Err:
#If DebugMode Then
540   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", MoveControl", mconModuleName
#End If
550   Resume MoveControl_Exit
End Function
'*****************************************************
Public Function MoveSplitter(IdSplitter As Long, MoveTo As Long) As Boolean   '4Matz:Changed
Attribute MoveSplitter.VB_Description = "Moves a splitter to the specified x- or y- (depending on the splitter's Orientation property) coordinate"
'*****************************************************
   ' Purpose    - Moves a splitter to the specified x- or y- (depending on the
   '              splitter's Orientation property) coordinate
   ' Effects    - * If successful, the control has been moved and all other
   '                effected splitters and controls' minimum and maximum x- and y-
   '                coordinates have been adjusted
   '              * If splitter IdSplitter doesn't exist, a run-time error has been
   '                generated
   ' Inputs     - * IdSplitter (a value that uniquely identifies the splitter the
   '                            developer want to move)
   '              * MoveTo (an integer value that specifies the x- or y- coordinate
   '                        (depending on the splitter's Orientation property)
   '                        where the splitter will be moved)
   '** 01/15/07 Yorgi- Performance & cleanup
   Dim lId                 As Long               'to determines the new friend control for the splitter
   Dim lId1                As Long
   Dim lId2                As Long
   Dim lTemp               As Long
   Dim oid1                As clsId              'for enumerating all Id in Ids collection
   Dim oid2                As clsId              'for enumerating all Id in Ids collection
10   On Error GoTo MoveSplitter_Err
     'TraceCtl  "MoveSplitter IdSplitter:" & IdSplitter & ", MoveTo:" & MoveTo
20   If Not mSplitters.IsExist(IdSplitter) Then
30      SecureRaiseError errIdSplitter, "MoveSplitter"
40   Else
50      With mSplitters(IdSplitter)
           '** 01/10/07 Yorgi- AutoCorrect Invalid MoveTos instead of raise error
           '-- If the destination coordinate is beyond the splitter's minimum or
           '   maximum value, reset MoveTo equal to Min and Max
60         If (MoveTo < .MinYc) Then MoveTo = .MinYc
70         If (MoveTo > .MaxYc) Then MoveTo = .MaxYc
80         Select Case .Orientation
              Case orHorizontal
                 '-- Move the splitter
90               .Yc = MoveTo
                 '-- Resize the controls and splitters effected by the splitter movement
100               For Each oid1 In .IdsCtlTop
110                  mVirtualControls.ItemNo(oid1.Id).Bottom = .Top
120                  Next
130               For Each oid1 In .IdsCtlBottom
140                  mVirtualControls.ItemNo(oid1.Id).Top = .Bottom
150                  Next
160               For Each oid1 In .IdsSplTop
170                  mSplitters(oid1.Id).Bottom = .Top
180                  Next
190               For Each oid1 In .IdsSplBottom
200                  mSplitters(oid1.Id).Top = .Bottom
210                  Next
                  '-- Finalizes the splitter movement by adjusting the minimum and
                  '   maximum y- coordinates of the splitters above or below the active splitter
220               If Not mblnDragSplitter Then
230                  For Each oid1 In .IdsCtlTop
240                     lId1 = oid1.Id            'localize values to avoid too many lookups
250                     If mVirtualControls.ItemNo(lId1).IdSplTop <> gconUninitializedLong Then
260                        lId = gconUninitializedLong
270                        With mSplitters(mVirtualControls.ItemNo(lId1).IdSplTop)
280                           For Each oid2 In .IdsCtlBottom
290                              lId2 = oid2.Id   'localize values to avoid too many lookups
300                              If lId = gconUninitializedLong Then
310                                 lId = lId2
320                              ElseIf mVirtualControls.ItemNo(lId2).Height - mVirtualControls.ItemNo(lId2).MinHeight < mVirtualControls.ItemNo(lId).Height - mVirtualControls.ItemNo(lId).MinHeight Then 'NOT lId...
330                                 lId = lId2
340                                 End If
350                              Next
360                           With mVirtualControls.ItemNo(lId)
370                              lTemp = .Bottom - .MinHeight
380                              End With
390                           .MaxYc = lTemp
400                           .IdCtlFriendBottom = lId
410                           End With
420                        End If
430                     Next
440                  For Each oid1 In .IdsCtlBottom
450                     lId1 = oid1.Id            'localize values to avoid too many lookups
460                     If mVirtualControls.ItemNo(lId1).IdSplBottom <> gconUninitializedLong Then
470                        lId = gconUninitializedLong
480                        With mSplitters(mVirtualControls.ItemNo(lId1).IdSplBottom)
490                           For Each oid2 In .IdsCtlTop
500                              lId2 = oid2.Id   'localize values to avoid too many lookups
510                              If lId = gconUninitializedLong Then
520                                 lId = lId2
530                              ElseIf mVirtualControls.ItemNo(lId2).Height - mVirtualControls.ItemNo(lId2).MinHeight < mVirtualControls.ItemNo(lId).Height - mVirtualControls.ItemNo(lId).MinHeight Then 'NOT lId...
540                                 lId = lId2
550                                 End If
560                              Next
570                           With mVirtualControls.ItemNo(lId)
580                              lTemp = .Top + .MinHeight
590                              End With
600                           .MinYc = lTemp
610                           .IdCtlFriendTop = lId
620                           End With
630                        End If
640                     Next
650                  End If
660            Case orVertical
670               .Xc = MoveTo                    ' Move the splitter
                  '-- Resize the controls and splitters that effected by the splitter movement
680               For Each oid1 In .IdsCtlLeft
690                  mVirtualControls.ItemNo(oid1.Id).Right = .Left
700                  Next
710               For Each oid1 In .IdsCtlRight
720                  mVirtualControls.ItemNo(oid1.Id).Left = .Right
730                  Next
740               For Each oid1 In .IdsSplLeft
750                  mSplitters(oid1.Id).Right = .Left
760                  Next
770               For Each oid1 In .IdsSplRight
780                  mSplitters(oid1.Id).Left = .Right
790                  Next
                  '-- Finalizes the splitter movement by adjusting the minimum and
                  '   maximum x- coordinates of the splitters above or below the active
                  '   splitter
800               If Not mblnDragSplitter Then
810                  For Each oid1 In .IdsCtlLeft
820                     lId1 = oid1.Id            'localize values to avoid too many lookups
830                     If mVirtualControls.ItemNo(lId1).IdSplLeft <> gconUninitializedLong Then
840                        lId = gconUninitializedLong
850                        With mSplitters(mVirtualControls.ItemNo(lId1).IdSplLeft)
860                           For Each oid2 In .IdsCtlRight
870                              lId2 = oid2.Id   'localize values to avoid too many lookups
880                              If lId = gconUninitializedLong Then
890                                 lId = lId2
900                              ElseIf mVirtualControls.ItemNo(lId2).Width - mVirtualControls.ItemNo(lId2).MinWidth < mVirtualControls.ItemNo(lId).Width - mVirtualControls.ItemNo(lId).MinWidth Then 'NOT lId...
910                                 lId = lId2
920                                 End If
930                              Next
940                           With mVirtualControls.ItemNo(lId)
950                              lTemp = .Right - .MinWidth
960                              End With
970                           .MaxXc = lTemp
980                           .IdCtlFriendRight = lId
990                           End With
1000                        End If
1010                     Next
1020                  For Each oid1 In .IdsCtlRight
1030                     lId1 = oid1.Id           'localize values to avoid too many lookups
1040                     If mVirtualControls.ItemNo(lId1).IdSplRight <> gconUninitializedLong Then
1050                        lId = gconUninitializedLong
1060                        With mSplitters(mVirtualControls.ItemNo(lId1).IdSplRight)
1070                           For Each oid2 In .IdsCtlLeft
1080                              lId2 = oid2.Id  'localize values to avoid too many lookups
1090                              If lId = gconUninitializedLong Then
1100                                 lId = lId2
1110                              ElseIf mVirtualControls.ItemNo(lId2).Width - mVirtualControls.ItemNo(lId2).MinWidth < mVirtualControls.ItemNo(lId).Width - mVirtualControls.ItemNo(lId).MinWidth Then 'NOT lId...
1120                                 lId = lId2
1130                                 End If
1140                              Next
1150                           With mVirtualControls.ItemNo(lId)
1160                              lTemp = .Left + .MinWidth
1170                              End With
1180                           .MinXc = lTemp
1190                           .IdCtlFriendLeft = lId
1200                           End With
1210                        End If
1220                     Next
1230                  End If
1240            End Select
1250         End With
1260      MoveSplitter = Refresh
1270      End If
1280   MoveSplitter_Exit:
1290   On Error GoTo 0
1300   Exit Function
1310   MoveSplitter_Err:
#If DebugMode Then
1320   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", MoveSplitter", mconModuleName
#End If
1330   Resume MoveSplitter_Exit
End Function
'*****************************************************
Private Sub mSplitters_BackColorChange(ByVal IdSplitter As Long)
'*****************************************************
   ' Purpose    - Refreshes the control back color to match the splitter's back
   '              color change
   ' Input      - IdSplitter (a value that uniquely identifies a splitter)
10   If IdSplitter <> gconUninitializedLong Then picSplitter(IdSplitter).BackColor = mSplitters(IdSplitter).BackColor
End Sub
'*****************************************************
Private Sub mSplitters_EnableChange(ByVal IdSplitter As Long)
'*****************************************************
   ' Purpose    - Refreshes the splitter back color to match the new property value
   ' Input      - IdSplitter (a value that uniquely identifies a splitter)
10   If IdSplitter <> gconUninitializedLong Then picSplitter(IdSplitter).Enabled = mSplitters(IdSplitter).Enable
End Sub
'*****************************************************
Private Sub mVirtualControls_TitleBarCloseVisibleChange(sIdControl As String)
'*****************************************************
   ' Purpose    - Refreshes the control's title bar close button visibility
   ' Input      - IdControl (a value that uniquely identifies a control)
10   With mVirtualControls(sIdControl)
20      ctbTitlebar(.TbarIdx).CloseVisible = .TitleBar_CloseVisible
30      End With
End Sub
'*****************************************************
Private Sub mVirtualControls_TitleBarTypeChange(sIdControl As String)
'*****************************************************
   ' Purpose    - Refreshes the control's title bar type
   ' Input      - IdControl (a value that uniquely identifies a control)
10   With mVirtualControls(sIdControl)
20      ctbTitlebar(.TbarIdx).TBarType = .TitleBar_TBarType
30      End With
End Sub
'*****************************************************
Private Sub mVirtualControls_TitleBarVisibleChange(sIdControl As String) '4Matz:Changed
'*****************************************************
   ' Purpose    - Refreshes the control's title bar visibility
   ' Input      - IdControl (a value that uniquely identifies a control)
   ' Effects    - The maximum and minimum value of the corresponding splitters have been adjusted
   '** 01/15/07 Yorgi- Performance & cleanup
   Dim oSpltr              As clsSplitter
10   If LenB(sIdControl) Then
20      With mVirtualControls(sIdControl)
30         If .TitleBar_Visible Then
40            If .IdSplTop <> gconUninitializedLong Then
50               Set oSpltr = mSplitters(.IdSplTop)
60               oSpltr.MaxYc = oSpltr.MaxYc - .TitleBar_VisibleHeight
70               End If
80            If .IdSplBottom <> gconUninitializedLong Then
90               Set oSpltr = mSplitters(.IdSplBottom)
100               oSpltr.MinYc = oSpltr.MinYc + .TitleBar_VisibleHeight
110               End If
120         Else
130            If .IdSplTop <> gconUninitializedLong Then
140               Set oSpltr = mSplitters(.IdSplTop)
150               oSpltr.MaxYc = oSpltr.MaxYc + .TitleBar_VisibleHeight
160               End If
170            If .IdSplBottom <> gconUninitializedLong Then
180               Set oSpltr = mSplitters(.IdSplBottom)
190               oSpltr.MinYc = oSpltr.MinYc - .TitleBar_VisibleHeight
200               End If
210            End If
220         End With
230      If Ambient.UserMode = False Then Refresh 'only during design time
240      End If
End Sub
'*****************************************************
Private Sub picSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '4Matz:New
'*****************************************************
   ' Purpose    - Handles picSlider MouseDown event if control alignment is active
   Dim mAlign              As Long
   Dim mlngContainerHwnd   As Long
   Dim rctUserControl      As RECT
10   mlngContainerHwnd = GetParentHwnd
20   GetWindowRect mlngContainerHwnd, mrctSlideArea
     'TraceCtl  "picSlider_MouseDown mrctSlideArea Left:" & mrctSlideArea.Left & ", Top:" & mrctSlideArea.Top & ", Bottom:" & mrctSlideArea.Bottom & ", Right:" & mrctSlideArea.Right
30   mAlign = Extender.Align
40   If mAlign Then                               'check for an aligned control
50      GetWindowRect UserControl.hWnd, rctUserControl
60      Select Case mAlign
           Case vbAlignLeft
70            oSlider.Position = picSlider.Left \ Screen.TwipsPerPixelX 'set the slider position
80            mrctSlideArea.Left = rctUserControl.Left 'adjust the slide window
90         Case vbAlignRight
100            oSlider.Position = mrctSlideArea.Right - mrctSlideArea.Left 'set the slider position
110            mrctSlideArea.Right = rctUserControl.Right 'adjust the slide window
120         End Select
         'TraceCtl  "picSlider_MouseDown oSlider.Position:" & oSlider.Position
130      End If
140   If oSlider.Orientation = cSPLTOrientationVertical Then
150      oSlider.SliderSize = picSlider.ScaleWidth ' set the slider width
160   Else
170      oSlider.Position = picSlider.Top \ Screen.TwipsPerPixelY
180      oSlider.SliderSize = picSlider.ScaleHeight 'set the slider height
190      End If
200   oSlider.MouseDown Button, mrctSlideArea.Top, mrctSlideArea.Left, mrctSlideArea.Bottom, mrctSlideArea.Right
End Sub
'*****************************************************
Private Sub picSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) '4Matz:New
'*****************************************************
   ' Purpose    - Handles picSlider MouseMove event if control alignment is active
10   oSlider.MouseMove Button, Shift, X, Y
End Sub
'*****************************************************
Private Sub picSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) '4Matz:New
'*****************************************************
   ' Purpose    - Handles picSlider MouseUp event if control alignment is active
   Dim lNewPos             As Long
   Dim lWidth              As Long
10   oSlider.MouseUp Button, Shift, X, Y
20   Select Case Extender.Align
        Case vbAlignLeft
30         lNewPos = oSlider.Position * Screen.TwipsPerPixelX
40      Case vbAlignRight
50         lNewPos = UserControl.Width - (oSlider.Delta * Screen.TwipsPerPixelX)
60      End Select
70   If oSlider.Orientation = cSPLTOrientationVertical Then
80      If lNewPos < picSlider.Width Then
90         lNewPos = picSlider.Width
100      Else
110         lWidth = (mrctSlideArea.Right - mrctSlideArea.Left) * Screen.TwipsPerPixelX
120         If lNewPos > lWidth Then lNewPos = lWidth
130         End If
140      UserControl.Width = lNewPos
150      End If
End Sub
'*****************************************************
Private Sub picSlider_Paint()
'*****************************************************
   Dim rec                 As RECT
10   picSlider.Cls
20   mdlAPI.SetRect rec, 0, 0, picSlider.ScaleWidth, picSlider.ScaleHeight
30   DrawEdge picSlider.hdc, rec, mdlAPI.EDGE_RAISED, mdlAPI.BF_RECT
End Sub
'*****************************************************
Private Sub picSplitter_Click(Index As Integer)
'*****************************************************
   ' Purpose    - Raises custom event SplitterClick
   ' Input      - Index
10   RaiseEvent SplitterClick(Index)
End Sub
'*****************************************************
Private Sub picSplitter_DblClick(Index As Integer)
'*****************************************************
   ' Purpose    - Raises custom event SplitterDblClick
   ' Input      - Index
10   RaiseEvent SplitterDblClick(Index)
End Sub
'*****************************************************
Private Sub picSplitter_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) '4Matz:Changed
'*****************************************************
   ' Purpose    - Initializes all things needed to move the splitter at run-time
   '              and raises custom event SplitterMouseDown and SplitterMoveBegin
   ' Assumption - Picture Box control picSplitter(Index) which represents the
   '              splitter exits
   ' Effects    - * mblnDrag = true
   '              * mlngDragStart = x or y (see the codes)
   '              * Control picSplitter(Index) is in front of the other controls
   '              * If the splitter's LiveUpdate property is false, then the
   '                picSpliter(Index) BackColor property has been set to the
   '                splitter's ActiveColor property
   '              * If the splitter's ClipCursor property is true, then the mouse
   '                pointer has been confined based on the splitter's MinXc, MinYc,
   '                MaxXc and MaxYc property value
   '              * Custom event SplitterMouseDown has been raised
   '              * If the user presses the left-button, then the SplitterMoveBegin
   '                event has been raised
   ' Inputs     - Index, Button, Shift, X, Y
   ' Note       - Notes that this procedure may confine the mouse pointer to
   '              certain area in the screen. If you call this procedure, don't
   '              forget to free the mouse pointer afterwards using
   '              mdlAPI.ClipCursorClear function.
   ' confine the mouse pointer
   '         pointer would be confined
   Dim oSplitter           As clsSplitter
   Dim uposCursor          As POINTAPI           'another variable needed to
   Dim urecClipCursor      As RECT               'the rectangle area where the mouse
10   If Button = vbLeftButton Then
        ' This subclassing below is used to handle the possibility of the user
        '   swithing to another application while dragging the splitter
20      mScIndex = Index
30      mScButton = Button
40      mScShift = Shift
50      mScX = X
60      mScY = Y
70      oSub.AddMsg mlngHwndRoot, WM_ACTIVATE, MSG_AFTER
80      mblnDragSplitter = True
90      Set oSplitter = mSplitters(CLng(Index))
100      Select Case oSplitter.Orientation
            Case orHorizontal
110            mlngDragStart = Y
120         Case orVertical
130            mlngDragStart = X
140         End Select
150      picSplitter(Index).ZOrder
160      If Not oSplitter.LiveUpdate Then
170         picSplitter(Index).BackColor = oSplitter.ActiveColor
180         UserControl.BackColor = oSplitter.BackColor
190         End If
200      If oSplitter.ClipCursor Then
210         mdlAPI.GetCursorPos uposCursor
220         uposCursor.X = (uposCursor.X * Screen.TwipsPerPixelX) - (picSplitter(Index).Left + X)
230         uposCursor.Y = (uposCursor.Y * Screen.TwipsPerPixelY) - (picSplitter(Index).Top + Y)
240         With urecClipCursor
250            Select Case oSplitter.Orientation
                  Case orHorizontal
260                  .Top = (uposCursor.Y + oSplitter.MinYc) \ Screen.TwipsPerPixelY
270                  .Right = (uposCursor.X + oSplitter.Right) \ Screen.TwipsPerPixelX
280                  .Bottom = (uposCursor.Y + oSplitter.MaxYc) \ Screen.TwipsPerPixelY
290                  .Left = (uposCursor.X + oSplitter.Left) \ Screen.TwipsPerPixelX
300               Case orVertical
310                  .Top = (uposCursor.Y + oSplitter.Top) \ Screen.TwipsPerPixelY
320                  .Right = (uposCursor.X + oSplitter.MaxXc) \ Screen.TwipsPerPixelX
330                  .Bottom = (uposCursor.Y + oSplitter.Bottom) \ Screen.TwipsPerPixelY
340                  .Left = (uposCursor.X + oSplitter.MinXc) \ Screen.TwipsPerPixelX
350               End Select
360            End With
370         mdlAPI.ClipCursor urecClipCursor
380         End If
390      RaiseEvent SplitterMoveBegin(Index, Shift, X, Y)
400      End If
410   muposPrev.X = X
420   muposPrev.Y = Y
430   RaiseEvent SplitterMouseDown(Index, Button, Shift, X, Y)
End Sub
'*****************************************************
Private Sub picSplitter_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute picSplitter_MouseMove.VB_Description = "Moves the splitter at run time"
'*****************************************************
   ' Purpose    - Moves the splitter at run-time and raises custom event
   '              SplitterMouseMove or SplitterMove
   ' Assumption - The picSplitter_MouseDown procedure has been called
   ' Effects    - * If the user moves the splitter, custom event Moving has been
   '                raised
   '              * Otherwise, custom event MouseMove has been raised
   '              * Other effect, as specified
   ' Inputs     - lIdx, Button, Shift, x, y
   Dim blnSplitterMoved    As Boolean            'indicating whether the splitter is moved
   Dim lIdx                As Long
   Dim lngPos              As Long               'to determine where the splitter will be moved
   Dim oSplitBar           As PictureBox
10   lIdx = CLng(Index)
20   With mSplitters(lIdx)
30      Select Case .Orientation
           Case orHorizontal
40            blnSplitterMoved = mblnDragSplitter And (Y <> mlngDragStart)
50         Case orVertical
60            blnSplitterMoved = mblnDragSplitter And (X <> mlngDragStart)
70         End Select
80      If blnSplitterMoved Then
90         Set oSplitBar = picSplitter(lIdx)
100         mblnSplitterMoved = True
110         Select Case .Orientation
               Case orHorizontal
120               lngPos = oSplitBar.Top + (Y - mlngDragStart)
130               If (lngPos < .MinYc) Then
140                  lngPos = .MinYc
150               ElseIf (lngPos + oSplitBar.Height > .MaxYc) Then
160                  lngPos = .MaxYc - oSplitBar.Height
170                  End If
180               oSplitBar.Top = lngPos
190               If .LiveUpdate Then MoveSplitter lIdx, oSplitBar.Top + (oSplitBar.Height \ 2)
200            Case orVertical
210               lngPos = oSplitBar.Left + (X - mlngDragStart)
220               If (lngPos < .MinXc) Then
230                  lngPos = .MinXc
240               ElseIf (lngPos + oSplitBar.Width > .MaxXc) Then
250                  lngPos = .MaxXc - oSplitBar.Width
260                  End If
270               oSplitBar.Left = lngPos
280               If .LiveUpdate Then MoveSplitter lIdx, oSplitBar.Left + (oSplitBar.Width \ 2)
290            End Select
300         End If
310      If Not mblnDragSplitter And Not blnSplitterMoved And ((X <> muposPrev.X) Or (Y <> muposPrev.Y)) Then
320         RaiseEvent SplitterMouseMove(lIdx, Button, Shift, X, Y)
330      ElseIf blnSplitterMoved Then
340         RaiseEvent SplitterMove(lIdx, Shift, X, Y)
350         End If
360      End With
370   muposPrev.X = X
380   muposPrev.Y = Y
End Sub
'*****************************************************
Private Sub picSplitter_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single) '4Matz:Changed
Attribute picSplitter_MouseUp.VB_Description = "Ends the run time splitter move action"
'*****************************************************
   ' Purpose    - Ends the run-time splitter move action and raises custom event
   '              SplitterMouseUp or SplitterMoveEnd
   ' Assumption - Picture Box control picSplitter(lIdx) which represents the
   '              splitter exits
   ' Effects    - * mblnDrag = false
   '              * Control picSplitter(lIdx) is in front of the other controls
   '              * If the splitter's LiveUpdate property is false, then the
   '                picSpliter(lIdx) BackColor property has been set to the
   '                splitter's BackColor property
   '              * The splitters minimum and maximum x- and y- coordinates have
   '                been adjusted
   '              * If the splitter's ClipCursor property is true, then the mouse
   '                pointer has been freed from confinement
   '              * If the splitter was moved then custom event Moved has been
   '                raised, otherwise, custom event MouseUp has been raised
   ' Inputs     - Index, Button, Shift, x, y
   Dim lIdx                As Long
10   oSub.DelMsg mlngHwndRoot, WM_ACTIVATE, MSG_AFTER
20   mblnDragSplitter = False
30   lIdx = CLng(Index)
     'TraceCtl  "picSplitter_MouseUp lIdx:" & lIdx
40   With picSplitter(lIdx)
50      If Not mSplitters(CLng(lIdx)).LiveUpdate Then .BackColor = mSplitters.BackColor
60      Select Case mSplitters(CLng(lIdx)).Orientation
           Case orHorizontal
70            MoveSplitter lIdx, .Top + (.Height \ 2)
80         Case orVertical
90            MoveSplitter lIdx, .Left + (.Width \ 2)
100         End Select
110      End With
120   If mSplitters(CLng(lIdx)).ClipCursor Then mdlAPI.ClipCursorClear
130   If mblnSplitterMoved Then
140      RaiseEvent SplitterMoveEnd(lIdx, Shift, X, Y)
150      mblnSplitterMoved = False
160   Else
170      RaiseEvent SplitterMouseUp(lIdx, Button, Shift, X, Y)
180      End If
End Sub
'*****************************************************
Private Function Refresh() As Boolean '4Matz:Changed
'*****************************************************
   ' Purpose    - Applies the virtual controls and splitters to their real controls and splitter
   '** 01/11/07 Yorgi- Add TBar Vertical functionality
   '** 02/04/07 Yorgi- Refresh is now re-entrant, allows rebuild later on errors (ie Usercontrol too small to draw componenets)
   Dim bTBarVisible        As Boolean
   Dim lhBrush             As Long
   Dim lHeight             As Long
   Dim lId                 As Long
   Dim lLeft               As Long
   Dim lngErrNumber        As Long               'for the control with r/o height
   Dim lngHeight           As Long               'adjusted height for list box control
   Dim lTBarHeight         As Long
   Dim lTop                As Long
   Dim lWidth              As Long
   Dim octl                As Control
   Dim ospl                As clsSplitter        'virtual control enumerator for Splitters collection
   Dim oVirtCtrl           As clsControl         'virtual control enumerator for Controls collection
   Dim rctBarStripe        As RECT
10   Const conErrHeightReadOnly As Long = 383
20   If mblnRefreshInProgress Then
        'TraceCtl  "<<<< Refresh InProgress >>>>"
30      Exit Function
40      End If
50   mblnRefreshInProgress = True
60   mblnLastRefreshOK = False
70   On Error GoTo Refresh_Err
     'TraceCtl  "Refresh Totals: mVirtualControls=" & mVirtualControls.Count & ", mSplitters=" & mSplitters.Count
     '-- Applies all virtuals splitters to their real splitters
80   For Each ospl In mSplitters
90      With picSplitter(ospl)
           'TraceCtl  "...ospl(" & ospl.Id & ") Left:" & ospl.Left & " Top:" & ospl.Top & " Width:" & ospl.Width & " Height:" & ospl.Height
100         .Move ospl.Left, ospl.Top, ospl.Width, ospl.Height
110         .BackColor = ospl.BackColor
120         .Enabled = ospl.Enable
130         .ZOrder
140         End With
150      Next
      '-- Applies all virtuals controls and its title bar to their real controls
      ''DumpCtls "Refresh"
160   For Each oVirtCtrl In mVirtualControls
170      If Not oVirtCtrl.Closed Then
180         lId = oVirtCtrl.TbarIdx
190         Set octl = oVirtCtrl.refCtlObj
200         lngHeight = AdjustedHeight(octl, oVirtCtrl)
210         With octl
               'move the TitleBar
220            With ctbTitlebar(lId)
230               lLeft = oVirtCtrl.Left
240               lTop = oVirtCtrl.Top
250               bTBarVisible = oVirtCtrl.TitleBar_Visible
260               .Visible = bTBarVisible
                  'only need to move titlebar if it is visible
270               If bTBarVisible Then
280                  lTBarHeight = oVirtCtrl.TitleBar_Height
290                  If .Orientation = TBO_HORIZONTAL Then 'Horizontal TitleBar
300                     lWidth = oVirtCtrl.Width
                        'TraceCtl  "...Move TBar(" & lId & ") Left:" & lLeft & " Top:" & lTop & " Width:" & lWidth & " Height:" & lTBarHeight
310                     .Move lLeft, lTop, lWidth, lTBarHeight 'move the titlebar
320                     lTop = lTop + lTBarHeight
330                     lHeight = lngHeight - lTBarHeight
340                     If lHeight < 0 Then lHeight = 0
350                  Else                         'Vertical TitleBar
360                     lWidth = oVirtCtrl.Width - lTBarHeight
370                     If lWidth < 0 Then lWidth = 0
380                     lHeight = lngHeight
                        'TraceCtl  "...Move TBar(" & lId & ") Left:" & lLeft & " Top:" & lTop & " Width:" & lWidth & " Height:" & lHeight
390                     .Move lLeft, lTop, lWidth, lHeight 'move the titlebar
400                     lLeft = lLeft + lTBarHeight
410                     End If
420               Else
430                  lWidth = oVirtCtrl.Width
440                  lHeight = lngHeight
450                  End If
460               End With
               'move the contained control
               'TraceCtl  "...Move VBCtrl(" & octl.Name & ") Left:" & lLeft & " Top:" & lTop & " Width:" & lWidth & " Height:" & lHeight
470            .Move lLeft, lTop, lWidth, lHeight
480            If lngErrNumber = conErrHeightReadOnly Then
490               .Move oVirtCtrl.Left, oVirtCtrl.Top, oVirtCtrl.Width
500               lngErrNumber = 0
510               End If
520            End With
530         End If
540      Next
550   mblnLastRefreshOK = True
560   Refresh = True
570   Refresh_Exit:
580   On Error Resume Next
590   mblnRefreshInProgress = False
      'no matter what happens, restore Slider position
600   If mlngSliderThickness > 0 Then             'only if visible
610      lTop = 0
620      lWidth = mlngSliderThickness
630      lHeight = UserControl.ScaleHeight
640      Select Case Extender.Align
            Case vbAlignLeft
650            lLeft = UserControl.ScaleWidth - mlngSliderThickness
660            If lLeft < 0 Then lLeft = 0
670         Case vbAlignRight
680            lLeft = 0
690         End Select
700      picSlider.Move lLeft, lTop, lWidth, lHeight
710      picSlider.ZOrder
         'TraceCtl  "...picSlider Left:" & lLeft & " Top:" & lTop & " Width:" & lWidth & " Height:" & lHeight
720      End If
      'TraceCtl  "Refresh Exit"
730   On Error GoTo 0
740   Exit Function
750   Refresh_Err:
760   If Err.Number = conErrHeightReadOnly Then
770      lngErrNumber = Err.Number
780      Resume Next
790      End If
#If DebugMode = 1 Then
800   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", Refresh", mconModuleName
#End If
810   Resume Refresh_Exit
End Function
'*****************************************************
Private Sub SecureRaiseError(ByVal udeErrNumber As genmErrNumber, Optional strSource As String = vbNullString) '4Matz:Changed
'*****************************************************
   ' Purpose    - Securely raises custom error udeErrNumber by firstly ends the
   '              subclassing
   ' Assumptions- * Error message udeErrNumber exists in the resource file
   '              * Global variable gstrControlName has been initialized
   ' Inputs     - * udeErrNumber
   '              * strSource (the location in form ClassNaRoutinesName where
   '                the error occur
   '** 01/20/07 Yorgi- Must continue to recieve WM_SIZE messages, even after errors occur
   ''''10   oSub.DelMsg mlngHwndParent, WM_SIZE, MSG_AFTER
10   RaiseError udeErrNumber, strSource
End Sub
'*****************************************************
Public Sub SetAlignment(ByVal eAlign As AlignConstants) '4Matz:New
'*****************************************************
   ' Purpose    - Changes in Align property determine Slider's visibility
10   If Extender.Align <> eAlign Then
20      Extender.Visible = False
30      Extender.Align = eAlign
40      InitSlider eAlign
50      Refresh
60      Extender.Visible = True
70      End If
End Sub
'*****************************************************
Public Sub SetMargins(ByVal lngMarginTop As Long, ByVal lngMarginLeft As Long, ByVal lngMarginBottom As Long, ByVal lngMarginRight As Long) '4Matz:New
'*****************************************************
   ' Purpose    - Sets the margins of the ActiveX Control from its container
   '              Minimize impact of rebuilts by calling Activate only once
10   If mlngMarginTop <> lngMarginTop Then
20      mlngMarginTop = lngMarginTop
30      PropertyChanged mconMarginTop
40      End If
50   If mlngMarginLeft <> lngMarginLeft Then
60      mlngMarginLeft = lngMarginLeft
70      PropertyChanged mconMarginLeft
80      End If
90   If mlngMarginBottom <> lngMarginBottom Then
100      mlngMarginBottom = lngMarginBottom
110      PropertyChanged mconMarginBottom
120      End If
130   If mlngMarginRight <> lngMarginRight Then
140      mlngMarginRight = lngMarginRight
150      PropertyChanged mconMarginRight
160      End If
      ''''170   Activate
End Sub
'*****************************************************
Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552
'*****************************************************
10   dlgAbout.Show vbModal
20   Unload dlgAbout
30   Set dlgAbout = Nothing
End Sub
'*****************************************************
Public Sub ShowControl(sIdControl As String, ByVal bOpen As Boolean, Optional ByRef Success As Boolean, Optional MaintainSize As Boolean = False) '4Matz:New
Attribute ShowControl.VB_Description = "Closes (hides) a control"
'*****************************************************
   ' Purpose    - Closes (hides) a control
   ' Effects    - * If successful, the control has been closed
   '              * If control IdControl doesn't exist, a run-time error has been generated
   '              * otherwise, no effect
   ' Input      - IdControl (a value that uniquely identifies the control the developer want to close)
   ' Return     - Success (a returned value that determines whether the Close method is successful)
10   On Error GoTo ShowControl_Err
     'TraceCtl  "ShowControl(" & sIdControl & ") begin"
20   If mVirtualControls.IsExist(sIdControl) Then
30      With mVirtualControls(sIdControl)
           '-- Close the virtual control IdControl
40         .Closed = Not bOpen
50         .refCtlObj.Visible = bOpen
60         ctbTitlebar(.TbarIdx).Visible = bOpen
70         Success = VCtrlManager
80         End With
90      End If
100   ShowControl_Exit:
      'TraceCtl  "ShowControl(" & sIdControl & ") rcode:" & Success
110   On Error GoTo 0
120   Exit Sub
130   ShowControl_Err:
140   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", ShowControl", mconModuleName
150   Resume ShowControl_Exit
End Sub
'*****************************************************
Public Property Get Size() As Long
Attribute Size.VB_Description = "Returns/sets the size of all splitters"
'*****************************************************
   ' Purpose    - Returns the size of all splitters
10   Size = mSplitters.Size
End Property
'*****************************************************
Public Property Let Size(ByVal lngSize As Long) '4Matz:Changed
'*****************************************************
   ' Purpose    - Sets the size of all splitters
   ' Effects    - * If Size is smaller than the splitters' minimum size then the
   '                splitters' size has been set to their minimum size
   '              * If there is a control with size less than its minimum size
   '                then the error message has been raised
   '              * Otherwise, as specified
   ' Input      - lngSize (the new Size property value)
   '** 01/15/07 Yorgi- Performance & cleanup
   Dim blnNeedToShow       As Boolean
   Dim lngdSize            As Long
   Dim octl                As clsControl         'for enumerating all virtual controls
   Dim oId                 As clsId              'for enumerating all Id in Ids collection
   Dim ospl                As clsSplitter        'for enumerating all virtual splitters
   Dim ospl2               As clsSplitter        'for enumerating all virtual splitters
10   On Error GoTo 0
     'TraceCtl  "Size lngSize:" & lngSize
20   If lngSize < mSplitters.MinimumSize Then lngSize = mSplitters.MinimumSize
30   lngdSize = (lngSize - mSplitters.Size) \ 2   '-- YorgiPerf: do the math one time!!!
40   mSplitters.Size = lngSize
50   PropertyChanged mconSize
     '-- Refresh the splitter size
60   blnNeedToShow = (mVirtualControls.Count = 0) And Ambient.UserMode
70   For Each octl In mVirtualControls
80      If Not octl.Closed Then
90         If octl.Left <> mVirtualControls.Left Then octl.Left = octl.Left + lngdSize
100         If octl.Top <> mVirtualControls.Top Then octl.Top = octl.Top + lngdSize
110         If octl.Right <> mVirtualControls.Right Then octl.Right = octl.Right - lngdSize
120         If octl.Bottom <> mVirtualControls.Bottom Then octl.Bottom = octl.Bottom - lngdSize
130         End If
140      Next
150   For Each ospl In mSplitters
160      Select Case ospl.Orientation
            Case orHorizontal
170            ospl.Height = lngSize
               '-- Adjust the width of the splitter if necessary
180            If ospl.Left > 0 Then ospl.Left = ospl.Left + (lngdSize)
190            If ospl.Right < Extender.Width Then ospl.Right = ospl.Right - (lngdSize)
               '-- Adjust the minimum value of the splitter if necessary
200            For Each oId In ospl.IdsCtlTop
210               If mVirtualControls.ItemNo(oId.Id).IdSplTop <> gconUninitializedLong Then
220                  ospl.MinYc = ospl.MinYc + (lngdSize)
230                  Exit For                     'loop varying oid
240                  End If
250               Next
               '-- Adjust the maximum value of the splitter if necessary
260            For Each oId In ospl.IdsCtlBottom
270               If mVirtualControls.ItemNo(oId.Id).IdSplBottom <> gconUninitializedLong Then
280                  ospl.MaxYc = ospl.MaxYc - (lngdSize)
290                  Exit For                     'loop varying oid
300                  End If
310               Next
320         Case orVertical
330            ospl.Width = lngSize
               '-- Adjust the height of the splitter if necessary
340            If ospl.Top > 0 Then ospl.Top = ospl.Top + (lngdSize)
350            If ospl.Bottom < Extender.Height Then ospl.Bottom = ospl.Bottom - (lngdSize)
               '-- Adjust the minimum value of the splitter if necessary
360            For Each oId In ospl.IdsCtlLeft
370               If mVirtualControls.ItemNo(oId.Id).IdSplLeft <> gconUninitializedLong Then
380                  ospl.MinXc = ospl.MinXc + (lngdSize)
390                  Exit For                     'loop varying oid
400                  End If
410               Next
               '-- Adjust the maximum value of the splitter if necessary
420            For Each oId In ospl.IdsCtlRight
430               If mVirtualControls.ItemNo(oId.Id).IdSplRight <> gconUninitializedLong Then
440                  ospl.MaxXc = ospl.MaxXc - (lngdSize)
450                  Exit For                     'loop varying oid
460                  End If
470               Next
480         End Select
490      Next
500   If mVirtualControls.IsValid Then
510      If Ambient.UserMode = False Then Refresh 'only during design time
520   Else
530      RaiseError errResizeSplitter, "Size"
540      End If
550   If blnNeedToShow Then Extender.Visible = mblnVisibleSave
End Property
'*****************************************************
Public Property Get Splitters() As clsSplitters
Attribute Splitters.VB_Description = "Returns a collection whose elements represent each virtual splitter in a Control Manager object"
'*****************************************************
10   Set Splitters = mSplitters
End Property
'*****************************************************
Private Sub StretchFillContainer() '4Matz:Changed
Attribute StretchFillContainer.VB_Description = "Stretchs the controls and splitters to fill-up their container"
'*****************************************************
   ' Purpose    - Stretches the controls and splitters to fill-up their container
   '** 02/15/07 Yorgi- Added error handling to StretchFillContainer and checked for valid Splitter.IdCtlFriends
   Dim lIdCtl              As Long               'local var IdCtl
   Dim octl                As clsControl         'for enumerating all virtual controls
   Dim oId                 As clsId              'for enumerating all Id in Ids collection
   Dim ospl                As clsSplitter        'for enumerating all virtual splitters
   Dim sngXScale           As Single             'a valid x-coorindate's scale
   Dim sngYScale           As Single             'a valid y-coordinate's scale
10   On Error GoTo StretchFillContainer_Err
     'TraceCtl  "StretchFillContainer Begin"
20   GetUCInnerDimensions mtypUCInside
     '-- StretchFillContainer the virtual splitters
30   If HasStretched(sngXScale, sngYScale) Then
40      mSplitters.Width = mSplitters.Width * sngXScale
50      mSplitters.Height = mSplitters.Height * sngYScale
60      For Each ospl In mSplitters
70         With ospl
80            Select Case .Orientation
                 Case orHorizontal
90                  .Xc = CLng(.Xc * sngXScale)
100                  .Yc = CLng((.Top * sngYScale) + ((.Height * sngYScale) / 2))
110                  .Width = CLng(.Width * sngXScale)
120                  lIdCtl = ospl.IdCtlFriendTop
130                  If lIdCtl <> gconUninitializedLong Then
140                     Set octl = mVirtualControls.ItemNo(lIdCtl)
150                     .MinYc = CLng((octl.Top * sngYScale) + octl.MinHeight)
160                     End If
170                  lIdCtl = ospl.IdCtlFriendBottom
180                  If lIdCtl <> gconUninitializedLong Then
190                     Set octl = mVirtualControls.ItemNo(lIdCtl)
200                     .MaxYc = CLng((octl.Bottom * sngYScale) - octl.MinHeight)
210                     End If
220               Case orVertical
230                  .Xc = CLng((.Left * sngXScale) + ((.Width * sngXScale) / 2))
240                  .Yc = CLng(.Yc * sngYScale)
250                  .Height = CLng(.Height * sngYScale)
260                  lIdCtl = ospl.IdCtlFriendLeft
270                  If lIdCtl <> gconUninitializedLong Then
280                     Set octl = mVirtualControls.ItemNo(lIdCtl)
290                     .MinXc = CLng((octl.Left * sngXScale) + octl.MinWidth)
300                     End If
310                  lIdCtl = ospl.IdCtlFriendRight
320                  If lIdCtl <> gconUninitializedLong Then
330                     Set octl = mVirtualControls.ItemNo(lIdCtl)
340                     .MaxXc = CLng((octl.Right * sngXScale) - octl.MinWidth)
350                     End If
360               End Select
370            End With
380         Next
390      For Each ospl In mSplitters
400         Select Case ospl.Orientation
               Case orHorizontal
410               For Each oId In ospl.IdsSplTop
420                  mSplitters(oId.Id).Bottom = ospl.Top
430                  Next
440               For Each oId In ospl.IdsSplBottom
450                  mSplitters(oId.Id).Top = ospl.Bottom
460                  Next
470            Case orVertical
480               For Each oId In ospl.IdsSplLeft
490                  mSplitters(oId.Id).Right = ospl.Left
500                  Next
510               For Each oId In ospl.IdsSplRight
520                  mSplitters(oId.Id).Left = ospl.Right
530                  Next
540            End Select
550         Next
         '-- StretchFillContainer the virtual controls
560      mVirtualControls.Width = mVirtualControls.Width * sngXScale
570      mVirtualControls.Height = mVirtualControls.Height * sngYScale
580      For Each octl In mVirtualControls
590         If Not octl.Closed Then
600            If octl.IdSplTop = gconUninitializedLong Then
610               octl.Top = mVirtualControls.Top
620               If octl.IdSplLeft <> gconUninitializedLong Then mSplitters(octl.IdSplLeft).Top = octl.Top
630               If octl.IdSplRight <> gconUninitializedLong Then mSplitters(octl.IdSplRight).Top = octl.Top
640            Else
650               octl.Top = mSplitters(octl.IdSplTop).Bottom
660               End If
670            If octl.IdSplRight = gconUninitializedLong Then
680               octl.Right = mVirtualControls.Right
690               If octl.IdSplTop <> gconUninitializedLong Then mSplitters(octl.IdSplTop).Right = octl.Right
700               If octl.IdSplBottom <> gconUninitializedLong Then mSplitters(octl.IdSplBottom).Right = octl.Right
710            Else
720               octl.Right = mSplitters(octl.IdSplRight).Left
730               End If
740            If octl.IdSplBottom = gconUninitializedLong Then
750               octl.Bottom = mVirtualControls.Bottom
760               If octl.IdSplLeft <> gconUninitializedLong Then mSplitters(octl.IdSplLeft).Bottom = octl.Bottom
770               If octl.IdSplRight <> gconUninitializedLong Then mSplitters(octl.IdSplRight).Bottom = octl.Bottom
780            Else
790               octl.Bottom = mSplitters(octl.IdSplBottom).Top
800               End If
810            If octl.IdSplLeft = gconUninitializedLong Then
820               octl.Left = mVirtualControls.Left
830               If octl.IdSplTop <> gconUninitializedLong Then mSplitters(octl.IdSplTop).Left = octl.Left
840               If octl.IdSplBottom <> gconUninitializedLong Then mSplitters(octl.IdSplBottom).Left = octl.Left
850            Else
860               octl.Left = mSplitters(octl.IdSplLeft).Right
870               End If
880            End If
890         Next
900      End If
      '** 02/01/07 Yorgi- Always try to refresh, we may need to redraw a Slider!!!
910   Refresh
920   StretchFillContainer_Exit:
930   On Error Resume Next
      'TraceCtl  "StretchFillContainer Exit"
940   On Error GoTo 0
950   Exit Sub
960   StretchFillContainer_Err:
#If DebugMode Then
970   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", StretchFillContainer", mconModuleName
#End If
980   Resume StretchFillContainer_Exit
End Sub
'*****************************************************
Private Sub TBarCreate(ByRef lngIdx As Long, ByRef sOwner As String) '4Matz:New
'*****************************************************
   ' Purpose    - Loads a TBar object
10   On Error Resume Next
20   lngIdx = ctbTitlebar.UBound + 1              'next TBar
30   Load ctbTitlebar(lngIdx)                     'Creates the new ctlTitleBar control instances to represent the control's title bar
40   With ctbTitlebar(lngIdx)
50      .Visible = True
60      .Tag = sOwner
70      End With
     'TraceCtl  "..TBarCreate:" & lngIdx & ", sOwner:" & sOwner
80   On Error GoTo 0
End Sub
'*****************************************************
Private Sub TBarRemove(lngIdx As Long) '4Matz:New
'*****************************************************
   ' Purpose    - Removes a TBar object
10   On Error Resume Next                         'Handles Err365-Unable to unload within this context
20   ctbTitlebar(lngIdx).Visible = False
30   Unload ctbTitlebar(lngIdx)
     'TraceCtl  "..TBarRemove:" & lngIdx
40   On Error GoTo 0
End Sub
'*****************************************************
Public Property Let TitleBar_CloseVisible(ByVal blnTitleBar_CloseVisible As Boolean)
'*****************************************************
   ' Purpose    - Sets a value that determines whether a close button in all
   '              control title bars is visible
   ' Input      - blnTitleBar_CloseVisible (the new TitleBar_CloseVisible property
   '                                        value)
10   If mVirtualControls.TitleBar_CloseVisible <> blnTitleBar_CloseVisible Then
20      mVirtualControls.TitleBar_CloseVisible = blnTitleBar_CloseVisible
30      PropertyChanged mconTitleBar_CloseVisible
40      If Ambient.UserMode = False Then Refresh  'only during design time
50      End If
End Property
'*****************************************************
Public Property Get TitleBar_CloseVisible() As Boolean
Attribute TitleBar_CloseVisible.VB_Description = "Returns/sets a value that determines whether a close button in all control title bars is visible"
'*****************************************************
   ' Purpose    - Returns a value that determines whether a close button in all
   '              control title bars is visible
10   TitleBar_CloseVisible = mVirtualControls.TitleBar_CloseVisible
End Property
'*****************************************************
Public Property Get TitleBar_Height() As Long
Attribute TitleBar_Height.VB_Description = "Returns/sets the height of all control title bars"
'*****************************************************
   ' Purpose    - Returns the height of the visible part of all control title bars
10   TitleBar_Height = mVirtualControls.TitleBar_Height
End Property
'*****************************************************
Public Property Let TitleBar_Height(ByVal lngTitleBar_Height As Long)
'*****************************************************
   ' Purpose    - Sets the height of all control title bars
   ' Input      - lngTitleBar_Height (the new TitleBar_Height property value)
10   mVirtualControls.TitleBar_Height = lngTitleBar_Height
20   PropertyChanged mconTitleBar_Height
End Property
'*****************************************************
Public Property Let TitleBar_Position(ByVal lngTitleBar_Position As TBarOrientation)
'*****************************************************
10   If mVirtualControls.TitleBar_Position <> lngTitleBar_Position Then
20      mVirtualControls.TitleBar_Position = lngTitleBar_Position
30      PropertyChanged mconTitleBar_Position
40      VCtrlManager
50      End If
End Property
'*****************************************************
Public Property Get TitleBar_Position() As TBarOrientation
'*****************************************************
10   TitleBar_Position = mVirtualControls.TitleBar_Position
End Property
'*****************************************************
Public Property Get TitleBar_TBarType() As TBarTypes
'*****************************************************
   ' Purpose    - Returns the TBarType of the all control title bars
10   TitleBar_TBarType = mVirtualControls.TitleBar_TBarType
End Property
'*****************************************************
Public Property Let TitleBar_TBarType(ByVal lngTitleBar_TBarType As TBarTypes)
'*****************************************************
   ' Purpose    - Sets the TBarType of all control title bars
   ' Input      - lngTitleBar_TBarType
10   If mVirtualControls.TitleBar_TBarType <> lngTitleBar_TBarType Then
20      mVirtualControls.TitleBar_TBarType = lngTitleBar_TBarType
30      PropertyChanged mconTitleBar_TBarType
40      If Ambient.UserMode = False Then Refresh  'only during design time
50      End If
End Property
'*****************************************************
Public Property Let TitleBar_Visible(ByVal blnTitleBar_Visible As Boolean)
'*****************************************************
   ' Purpose    - Sets a value that determines whether all control title bars are visible
   ' Input      - blnblnTitleBar_Visible (the new TitleBar_Visible property value)
   Dim blnSuccess          As Boolean
   Dim lngd                As Long
   Dim octl                As clsControl         'for enumerating all virtual controls in Controls collection
   Dim ospl                As clsSplitter        'for enumerating all virtual splitters in Splitters collection
10   If mVirtualControls.TitleBar_Visible <> blnTitleBar_Visible Then
20      mVirtualControls.TitleBar_Visible = blnTitleBar_Visible
30      If blnTitleBar_Visible Then
40         lngd = mVirtualControls.TitleBar_VisibleHeight
50      Else
60         lngd = -mVirtualControls.TitleBar_VisibleHeight
70         End If
80      blnSuccess = True
90      For Each octl In mVirtualControls
100         If octl.TitleBar_Visible <> mVirtualControls.TitleBar_Visible Then
110            blnSuccess = False
120            Exit For                           'loop varying octl
130            End If
140         Next
150      If blnSuccess Then blnSuccess = blnSuccess And mVirtualControls.IsValid
160      If blnSuccess Then
170         PropertyChanged mconTitleBar_Visible
180         For Each ospl In mSplitters
190            If ospl.Orientation = orHorizontal Then
200               If ospl.IdsCtlTop.Count > 0 Then ospl.MinYc = ospl.MinYc + lngd
210               If ospl.IdsCtlBottom.Count > 0 Then ospl.MaxYc = ospl.MaxYc - lngd
220               End If
230            Next
240         If Ambient.UserMode = False Then Refresh 'only during design time
250      Else
260         mVirtualControls.TitleBar_Visible = Not mVirtualControls.TitleBar_Visible
270         End If
280      End If
End Property
'*****************************************************
Public Property Get TitleBar_Visible() As Boolean
Attribute TitleBar_Visible.VB_Description = "Returns/sets a value that determines whether all control title bars are visible"
'*****************************************************
   ' Purpose    - Returns a value that determines whether all control title bars are visible
10   TitleBar_Visible = mVirtualControls.TitleBar_Visible
End Property
'*****************************************************
Public Sub TraceCtl(sMsg As String)
'*****************************************************
#If DebugMode Then
10   AppTrace mconModuleName, gstrParentName, sMsg
#End If
End Sub
'*****************************************************
Public Function UnDock(df As DokNForm, Optional bRemove As Boolean) As Boolean '4Matz:Changed
'*****************************************************
   ' Purpose   - Undock and show the form
   Dim bSuccess            As Boolean            'local return status
   Dim ofrm                As Form
   Dim Style               As Long
   Dim sVCtlKey            As String
10   Const constSource As String = "UnDock"
20   On Error GoTo UnDock_Err
30   Set ofrm = df.DockedForm
40   If ofrm Is Nothing Then
50      bRemove = True                            ' if the form unloaded, destroy df object
60      End If
     'TraceCtl  constSource & " begin bRemove:" & bRemove
70   If (Not bRemove) And ((df.Style And DSFloat) = False) Then ' if this form can not float then exit
80      Exit Function
90      End If
100   sVCtlKey = df.VCtlKey                       'get the VirtCtl key stored in df
110   ofrm.Visible = False                        'hide the window to avoid flicker
120   SetParent ofrm.hWnd, df.FormParentHwnd      'ReStore our parent
130   If bRemove Then
140      HostCtrlRemove df.HostContainer          'remove Host Control from ContainedControls
150      VCtrlRemove df.VCtlKey                   'remove host control from the VirtCtls collection
160      VCtrlManager                             'refresh splitters & refresh
170   Else
         ' restore form's style
180      SetWindowLong ofrm.hWnd, GWL_STYLE, df.FloatingStyle
         ' restore form's extended style
190      SetWindowLong ofrm.hWnd, GWL_EXSTYLE, df.FloatingExStyle
         ' move window to its floating position
200      ofrm.Move df.FloatingLeft, df.FloatingTop, df.FloatingWidth, df.FloatingHeight
210      ofrm.Visible = True                      ' change visiblity
220      ofrm.ZOrder
230      df.State = DS_UnDocked
240      ShowControl sVCtlKey, False              'hide df and rebuild display
250      End If
260   Set ofrm = Nothing
270   UnDock = True
      'TraceCtl  constSource & " end"
280   UnDock_Exit:
290   On Error GoTo 0
300   Exit Function
310   UnDock_Err:
320   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", UnDock", mconModuleName
330   Resume UnDock_Exit
End Function
'*****************************************************
Public Property Let UnloadFrmOnClose(bln As Boolean) '4Matz:New
'*****************************************************
   ' Purpose    - Flags form for unloading if HostControl is closed
10   mblnUnloadFrmOnClose = bln
20   PropertyChanged mconUnloadFrmOnClose
End Property
'*****************************************************
Public Property Get UnloadFrmOnClose() As Boolean '4Matz:New
'*****************************************************
   ' Purpose    - Flags form for unloading if HostControl is closed
10   UnloadFrmOnClose = mblnUnloadFrmOnClose
End Property
'*****************************************************
Private Sub UserControl_Initialize() '4Matz:Changed
'*****************************************************
10   Set mVirtualControls = New clsControls
20   Set mSplitters = New clsSplitters
30   Set oDockedForms = New DokNForms
40   Set oSub = New cSubclass
50   Set oSlider = New clsSlider
     'TraceCtl  "UserControl_Initialize"
End Sub
'*****************************************************
Private Sub UserControl_InitProperties()
'*****************************************************
10   mblnFillContainer = mconDefaultFillContainer
20   mlngMarginBottom = mconDefaultMarginBottom
30   mlngMarginLeft = mconDefaultMarginLeft
40   mlngMarginRight = mconDefaultMarginRight
50   mlngMarginTop = mconDefaultMarginTop
60   mSplitters.BackColor = Ambient.BackColor
     'TraceCtl  "UserControl_InitProperties"
End Sub
'*****************************************************
Private Sub UserControl_ReadProperties(PropBag As PropertyBag) '4Matz:Changed
'*****************************************************
#If DebugMode Then
10   gstrParentName = UserControl.Parent.Name
#End If
20   With PropBag
30      mSplitters.ActiveColor = .ReadProperty(mconActiveColor, mSplitters.DefaultActiveColor)
40      mSplitters.BackColor = .ReadProperty(mconBackColor, Ambient.BackColor)
50      mSplitters.ClipCursor = .ReadProperty(mconClipCursor, mSplitters.DefaultClipCursor)
60      mSplitters.Enable = .ReadProperty(mconEnable, mSplitters.DefaultEnable)
70      mblnFillContainer = .ReadProperty(mconFillContainer, mconDefaultFillContainer)
80      mSplitters.LiveUpdate = .ReadProperty(mconLiveUpdate, mSplitters.DefaultLiveUpdate)
90      mlngMarginBottom = .ReadProperty(mconMarginBottom, mconDefaultMarginBottom)
100      mlngMarginLeft = .ReadProperty(mconMarginLeft, mconDefaultMarginLeft)
110      mlngMarginRight = .ReadProperty(mconMarginRight, mconDefaultMarginRight)
120      mlngMarginTop = .ReadProperty(mconMarginTop, mconDefaultMarginTop)
130      mSplitters.Size = .ReadProperty(mconSize, mSplitters.DefaultSize)
140      mVirtualControls.TitleBar_CloseVisible = .ReadProperty(mconTitleBar_CloseVisible, mVirtualControls.DefaultTitleBar_CloseVisible)
150      mVirtualControls.TitleBar_Height = .ReadProperty(mconTitleBar_Height, mVirtualControls.DefaultTitleBar_Height)
160      mVirtualControls.TitleBar_Visible = .ReadProperty(mconTitleBar_Visible, mVirtualControls.DefaultTitleBar_Visible)
170      mVirtualControls.TitleBar_TBarType = .ReadProperty(mconTitleBar_TBarType, mVirtualControls.DefaultTitleBar_TBarType)
180      mVirtualControls.TitleBar_Position = .ReadProperty(mconTitleBar_Position, mVirtualControls.DefaultTitleBar_Position)
190      mblnUnloadFrmOnClose = .ReadProperty(mconUnloadFrmOnClose, mconDefaultUnloadFrmOnClose)
200      End With
210   gstrControlName = Ambient.DisplayName
      'TraceCtl  "UserControl_ReadProperties Ambient.UserMode:" & Ambient.UserMode & ", mSplitters.Size:" & mSplitters.Size
      ' Hide the ActiveX control when initializing the controls in it to reduce the flickering
220   If Ambient.UserMode Then
230      mlngHwndParent = UserControl.Parent.hWnd
240      mblnVisibleSave = Extender.Visible
250      Extender.Visible = False
260      oSub.Subclass mlngHwndParent, Me
270      oSub.AddMsg mlngHwndParent, WM_SHOWWINDOW, MSG_AFTER
280      oSub.AddMsg mlngHwndParent, WM_SIZE, MSG_AFTER
290      End If
End Sub
'*****************************************************
Private Sub UserControl_Resize()
Attribute UserControl_Resize.VB_Description = "Adjusts the components inside the control to agree with the control's size"
'*****************************************************
   ' Purpose    - Adjusts the components inside the control to agree with the control's size
   ' Effect     - See the codes below and see the effects of VCtrlManager and StretchFillContainer procedures
    'TraceCtl  "UserControl_Resize"
10   GetWindowRect UserControl.hWnd, mrctUserControl 'get the Usercontrol window rect
20   If (Not (mVirtualControls.Count = ContainedControls.Count)) Then
30      VCtrlManager True                         'its safer to always rebuild in design mode
40   ElseIf mblnLastRefreshOK = False Then
50      VCtrlManager
60   Else
70      StretchFillContainer
80      End If
End Sub
'*****************************************************
Private Sub UserControl_Terminate() '4Matz:Changed
'*****************************************************
10   On Error Resume Next
     'TraceCtl  "UserControl_Terminate begin"
20   Set oSub = Nothing
30   Set oSlider = Nothing
     'TraceCtl  "UserControl_Terminate oDockedForms:" & oDockedForms.Count
40   Set oDockedForms = Nothing
     'TraceCtl  "UserControl_Terminate mVirtualControls:" & mVirtualControls.Count
50   Set mVirtualControls = Nothing
     'TraceCtl  "UserControl_Terminate mSplitters:" & mSplitters.Count
60   Set mSplitters = Nothing
     'TraceCtl  "UserControl_Terminate end"
End Sub
'*****************************************************
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'*****************************************************
10   With PropBag
20      .WriteProperty mconActiveColor, mSplitters.ActiveColor, mSplitters.DefaultActiveColor
30      .WriteProperty mconBackColor, mSplitters.BackColor, Ambient.BackColor
40      .WriteProperty mconClipCursor, mSplitters.ClipCursor, mSplitters.DefaultClipCursor
50      .WriteProperty mconEnable, mSplitters.Enable, mSplitters.DefaultEnable
60      .WriteProperty mconFillContainer, mblnFillContainer, mconDefaultFillContainer
70      .WriteProperty mconLiveUpdate, mSplitters.LiveUpdate, mSplitters.DefaultLiveUpdate
80      .WriteProperty mconMarginBottom, mlngMarginBottom, mconDefaultMarginBottom
90      .WriteProperty mconMarginLeft, mlngMarginLeft, mconDefaultMarginLeft
100      .WriteProperty mconMarginRight, mlngMarginRight, mconDefaultMarginRight
110      .WriteProperty mconMarginTop, mlngMarginTop, mconDefaultMarginTop
120      .WriteProperty mconSize, mSplitters.Size, mSplitters.DefaultSize
130      .WriteProperty mconTitleBar_CloseVisible, mVirtualControls.TitleBar_CloseVisible, mVirtualControls.DefaultTitleBar_CloseVisible
140      .WriteProperty mconTitleBar_Height, mVirtualControls.TitleBar_Height, mVirtualControls.DefaultTitleBar_Height
150      .WriteProperty mconTitleBar_Visible, mVirtualControls.TitleBar_Visible, mVirtualControls.DefaultTitleBar_Visible
160      .WriteProperty mconTitleBar_TBarType, mVirtualControls.TitleBar_TBarType, mVirtualControls.DefaultTitleBar_TBarType
170      .WriteProperty mconTitleBar_Position, mVirtualControls.TitleBar_Position, mVirtualControls.DefaultTitleBar_Position
180      .WriteProperty mconUnloadFrmOnClose, mblnUnloadFrmOnClose, mconDefaultUnloadFrmOnClose
190      End With
End Sub
'*****************************************************
Private Function VCtrlAdd(octl As Control, Optional oVirtCtl As clsControl, Optional sCtlName As String) As Long           '4Matz:New
'*****************************************************
   ' Purpose    - Creates a new TBar & mVirtualControls object to add to the collection
   ' Returns    - The new VirtCtl IdName & TBarIdx
   Dim lTbarIdx            As Long               'local var for indexing
   Dim sTitleBarCaption    As String             'user assignable caption for TBar
10   On Error GoTo VCtrlAdd_Err
     'TraceCtl  "VirtCntrlAdd ..cctl(" & mVirtualControls.Count + 1 & "):" & octl.Name
     '-- DoknSplitz control can't have another DoknSplitz control inside it
20   If TypeOf octl Is ControlManager Then SecureRaiseError errSelfContained, "Init"
     '-----------------------------------------
     '-- Create a new virtual control for the ContainedControl
     '-----------------------------------------
30   mVirtualControls.Add octl, oVirtCtl          'create oVirtCtl and add to the collection
     '-----------------------------------------
     '-- Build new ctlTitleBar control
     '-----------------------------------------
40   sCtlName = oVirtCtl.Key
50   TBarCreate lTbarIdx, sCtlName                'Creates the new ctlTitleBar control instances to represent the control's title bar
60   oVirtCtl.TbarIdx = lTbarIdx                  'store our TBar Index in the VirtCtl
70   With ctbTitlebar(lTbarIdx)
80      .TBarType = mVirtualControls.TitleBar_TBarType
90      .Orientation = mVirtualControls.TitleBar_Position
100      .CloseVisible = mVirtualControls.TitleBar_CloseVisible
         '-----------------------------------------
         'see if owner wants to init a caption
         '-----------------------------------------
110      If LenB(.Caption) = 0 Then
120         RaiseEvent TitleBarCaption(sCtlName, sTitleBarCaption)
130         If LenB(sTitleBarCaption) Then .Caption = sTitleBarCaption
140         End If
150      End With
160   mblnLastRefreshOK = False
170   VCtrlAdd = lTbarIdx
180   VCtrlAdd_Exit:
190   On Error Resume Next
      'TraceCtl  "VirtCntrlAdd End"
200   On Error GoTo 0
210   Exit Function
220   VCtrlAdd_Err:
230   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", VCtrlAdd", mconModuleName
240   Resume VCtrlAdd_Exit
End Function
'*****************************************************
Private Function VCtrlIdxToDoknForm(ByRef sVKey As String, ByRef df As DokNForm) As Boolean   '4Matz:New
'*****************************************************
   ' Purpose    - Get the docked form object
   Dim sDFKey              As String
10   If mVirtualControls.IsExist(sVKey) Then      'make sure we have a valid key
20      sDFKey = mVirtualControls(sVKey).DFKey
30      If LenB(sDFKey) Then
40         Set df = oDockedForms(sDFKey)
50         VCtrlIdxToDoknForm = Not (df Is Nothing)
60         End If
70      End If
End Function
'*****************************************************
Public Function VCtrlManager(Optional blnRebuildAll As Boolean = False) As Boolean  '4Matz:New
'*****************************************************
   ' Purpose    - Builds virtual controls and splitters, allowing Refresh function to
   '              apply virtual dimensions to the real controls and splitters
   ' Effect     - * If successed, as specified
   '              * Otherwise, the custom error message has been raised
   '** 01/26/07 Yorgi- Performance & cleanup
   Dim lngIdx              As Long               'local var for indexing
   Dim lTemp               As Long               'local var for long math
   Dim lTot                As Long               'local var for indexing
   Dim oId                 As clsId              'for enumerating all Id in Ids collection
   Dim ospl                As clsSplitter        'for enumerating all virtual splitters in Splitters collection
   Dim oVirtCtl            As clsControl         'for enumerating all virtual controls in Controls collection
10   On Error GoTo VCtrlManager_Err
     'TraceCtl  "VirtCntrlManager BatchBuild:" & mblnBatchBuild & ", NewControl:" & blnRebuildAll
20   If mblnBatchBuild Then Exit Function         'come back and process all pending later
30   GetUCInnerDimensions mtypUCInside
40   If blnRebuildAll Then
50      VCtrlRebuildAll                           'first time around build all the controls and TBars
60   ElseIf Not mblnLastRefreshOK Then
70      mVirtualControls.ReCalcAllDimensions      'refresh fails if sizes are too small so recalc dims
80      End If
     '-----------------------------------------
     '-- Make the virtual controls solid and fill-up the DoknSplitz control's container
     '-----------------------------------------
90   With mVirtualControls
100      .Left = mtypUCInside.Left
110      .Top = mtypUCInside.Top
120      .Right = mtypUCInside.Width
130      .Bottom = mtypUCInside.Height
140      .RemoveHeaps
150      .Compact
160      .RemoveHoles
170      .Stretch
180      End With
      '-----------------------------------------
      '-- Build virtual splitters and place it as the virtual controls' "border"
      '-----------------------------------------
190   For Each oVirtCtl In mVirtualControls
200      oVirtCtl.IdSplTop = gconUninitializedLong
210      oVirtCtl.IdSplRight = gconUninitializedLong
220      oVirtCtl.IdSplBottom = gconUninitializedLong
230      oVirtCtl.IdSplLeft = gconUninitializedLong
240      Next
      '-----------------------------------------
      '-- Create new splitter objects to manage picSplitters
      '-----------------------------------------
250   With mSplitters
260      .Left = mtypUCInside.Left
270      .Top = mtypUCInside.Top
280      .Right = mtypUCInside.Width
290      .Bottom = mtypUCInside.Height
300      .Clear
         '-- Create splitter objects only for open controls
310      lTot = mVirtualControls.Count
320      For lngIdx = 1 To lTot
330         Set oVirtCtl = mVirtualControls.ItemNo(lngIdx)
340         If Not oVirtCtl.Closed Then
350            .Add oVirtCtl, mVirtualControls, lngIdx
360            End If
370         Next
380      For Each ospl In mSplitters
390         ospl.IdsSplTop.RemoveDeleted .Count
400         ospl.IdsSplRight.RemoveDeleted .Count
410         ospl.IdsSplBottom.RemoveDeleted .Count
420         ospl.IdsSplLeft.RemoveDeleted .Count
430         Next
         '** 01/26/07 Yorgi- Performance & cleanup
440      For Each ospl In mSplitters
450         With ospl
               'TraceCtl  "..VirtCntrlManager Splitter(" & .Id & ") Top=" & .Top & ", Bottom=" & .Bottom & ", Left=" & .Left & ", Right=" & .Right
460            End With
470         Select Case ospl.Orientation
               Case orHorizontal
480               lTemp = (ospl.Height \ 2)       '-- YorgiPerf: do the math one time!!!
490               For Each oId In ospl.IdsSplTop
500                  With .Item(oId.Id)           '-- YorgiPerf: do the lookup one time!!!
510                     .Bottom = .Bottom - lTemp
520                     End With
530                  Next
540               For Each oId In ospl.IdsSplBottom
550                  With .Item(oId.Id)           '-- YorgiPerf: do the lookup one time!!!
560                     .Top = .Top + lTemp
570                     End With
580                  Next
590            Case orVertical
600               lTemp = (ospl.Width \ 2)        '-- YorgiPerf: do the math one time!!!
610               For Each oId In ospl.IdsSplLeft
620                  With .Item(oId.Id)           '-- YorgiPerf: do the lookup one time!!!
630                     .Right = .Right - lTemp
640                     End With
650                  Next
660               For Each oId In ospl.IdsSplRight
670                  With .Item(oId.Id)           '-- YorgiPerf: do the lookup one time!!!
680                     .Left = .Left + lTemp
690                     End With
700                  Next
710            End Select
720         Next
730      End With
      '-----------------------------------------
      '-- Remove existing picSplitter controls
      '-----------------------------------------
740   lTemp = picSplitter.Count - 2               ' always leave last occurance
750   On Error Resume Next                        'Handles Err365-Unable to unload within this context
760   For lngIdx = 0 To lTemp
770      picSplitter(lngIdx).Visible = False
780      Unload picSplitter(lngIdx)
790      Next
800   On Error GoTo VCtrlManager_Err
      '-----------------------------------------
      '-- Creates the new PictureBox control instances to represent the splitter
      '-----------------------------------------
810   For Each ospl In mSplitters
820      CreateSplitr ospl.Id                     '-- Creates the new PictureBox control instances to represent the splitter
830      With picSplitter(ospl)                   '-- YorgiPerf: do the lookup one time!!!
840         .MousePointer = vbCustom
850         Select Case ospl.Orientation
               Case orHorizontal
860               .MouseIcon = LoadResPicture(gconCurHSplitter, vbResCursor)
870            Case orVertical
880               .MouseIcon = LoadResPicture(gconCurVSplitter, vbResCursor)
890            End Select
900         .Visible = True
910         End With
920      Next
930   If IsSolid Then VCtrlManager = Refresh
940   VCtrlManager_Exit:
      'TraceCtl  "VirtCntrlManager End"
950   If Ambient.UserMode Then Extender.Visible = mblnVisibleSave
960   On Error GoTo 0
970   Exit Function
980   VCtrlManager_Err:
990   If Not Ambient.UserMode Then Resume Next
1000   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", VCtrlManager", mconModuleName
1010   Resume VCtrlManager_Exit
End Function
'*****************************************************
Private Sub VCtrlRebuildAll() '4Matz:New
'*****************************************************
   ' Purpose    - Builds a new mVirtualControls collection for each ContainedControl
   ' Effect     - * If successed, as specified
   '              * Otherwise, the custom error message has been raised
   Dim blnTBarCloseVisibleSave As Boolean
   Dim blnTBarVisibleSave  As Boolean
   Dim lngIdx              As Long               'local var for indexing
   Dim lngIdxTot           As Long               'local var for indexing to total count
   Dim lTitleBarPosSave    As TBarOrientation
   Dim lTitleBarTypeSave   As TBarTypes
   Dim octl                As Control
10   On Error GoTo VCtrlRebuildAll_Err
     'TraceCtl  "VirtCntrlRebuildAll mVirtualControls:" & mVirtualControls.Count & " ContainedControls:" & ContainedControls.Count
20   If mVirtualControls.Count Then
        '-----------------------------------------
        '-- Save the TitleBar properties
        '-----------------------------------------
30      lTitleBarPosSave = mVirtualControls.TitleBar_Position
40      blnTBarCloseVisibleSave = mVirtualControls.TitleBar_CloseVisible
50      blnTBarVisibleSave = mVirtualControls.TitleBar_Visible
60      lTitleBarTypeSave = (mVirtualControls.TitleBar_TBarType Or TBT_SINGLESTRIPE) And Not TBT_DOUBLESTRIPE
        '-----------------------------------------
        '-- Create new Controls collection
        '-----------------------------------------
70      Set mVirtualControls = New clsControls
        '-----------------------------------------
        '-- ReStore the TitleBar properties
        '-----------------------------------------
80      mVirtualControls.TitleBar_TBarType = lTitleBarTypeSave
90      mVirtualControls.TitleBar_CloseVisible = blnTBarCloseVisibleSave
100      mVirtualControls.TitleBar_Visible = blnTBarVisibleSave
110      mVirtualControls.TitleBar_Position = lTitleBarPosSave
120      End If
      '-- Unload existing ctlTitleBar control instances
130   lngIdxTot = ctbTitlebar.Count - 1           'ignore index 0 TBar, used as the seed
140   For lngIdx = 1 To lngIdxTot
150      TBarRemove lngIdx
160      Next
170   On Error GoTo VCtrlRebuildAll_Err
      '-----------------------------------------
      '-- Create a new virtual control for each ContainedControl
      '-----------------------------------------
180   lngIdxTot = ContainedControls.Count - 1
190   If lngIdxTot >= 0 Then
         'TraceCtl  "VirtCntrlRebuildAll ContainedCntrls:" & ContainedControls.Count
200      For Each octl In ContainedControls
210         VCtrlAdd octl
220         Next
230      End If
240   VCtrlRebuildAll_Exit:
250   On Error Resume Next
      'TraceCtl  "VirtCntrlRebuildAll End"
260   mblnLastRefreshOK = False
270   On Error GoTo 0
280   Exit Sub
290   VCtrlRebuildAll_Err:
300   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", VCtrlRebuildAll", mconModuleName
310   Resume VCtrlRebuildAll_Exit
End Sub
'*****************************************************
Private Function VCtrlRemove(sVCtlKey As String) '4Matz:New
'*****************************************************
   ' Purpose    - Removes TBar & mVirtualControls object
10   On Error GoTo VCtrlRemove_Err
     'TraceCtl  "VirtCntrlRemove ..vctl:" & sVCtlKey
20   With mVirtualControls(sVCtlKey)
30      TBarRemove .TbarIdx                       'Remove ctlTitleBar control
40      mVirtualControls.Remove .Key              'Remove virtual control from the mVirtualControls collection
50      End With
60   VCtrlRemove_Exit:
     'TraceCtl  "VirtCntrlRemove End"
70   On Error GoTo 0
80   Exit Function
90   VCtrlRemove_Err:
100   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", VCtrlRemove", mconModuleName
110   Resume VCtrlRemove_Exit
End Function

' Yorgi's 4Matz [Feb 28,2007 23:58:51] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

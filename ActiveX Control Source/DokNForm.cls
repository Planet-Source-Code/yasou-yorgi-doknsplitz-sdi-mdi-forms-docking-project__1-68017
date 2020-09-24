VERSION 1.0 CLASS
BEGIN
MultiUse = -1                                     'True
Persistable = 0                                   'NotPersistable
DataBindingBehavior = 0                           'vbNone
DataSourceBehavior  = 0                           'vbNone
MTSTransactionMode  = 0                           'NotAnMTSObject
END
Attribute VB_Name = "DokNForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Const m_constClassName         As String = "DokNForm"
Private Const g_ErrConstant            As Long = vbObjectError + 1000
Private Const ControlManager_HEIGHT    As Integer = 80
Private Const ControlManager_WIDTH     As Integer = 80
' Keep up with the errors
Private mlngErrNum                     As Long
Private m_strErrStr                    As String
Private m_strErrSource                 As String
' class properties
Private mlngMinWidth                   As Long
Private mlngMinHeight                  As Long
Private mblnOverMax                    As Boolean
Private mblnOverClose                  As Boolean
Private mlngParenthWnd                 As Long
Private mblnStyleChanged               As Boolean
Private m_TBarType                     As TBarTypes
Private m_TBarPos                      As TBarOrientation ' TBar Position/Orientation
Private mlngFloatingStyle              As Long    ' Floating style
Private mlngFloatingExStyle            As Long    ' Floating extended style
Private mlngFloatingWidth              As Long    ' Floating dimensions
Private mlngFloatingHeight             As Long
Private mlngFloatingLeft               As Long
Private mlngFloatingTop                As Long
Private m_State                        As eDockState ' form state dock/undocked
Private m_Style                        As eDockStyles ' availlable style
Private m_ctlHostContainer             As PictureBox ' picturebox as form's docked container
Private m_AttachToCtrlPtr              As Long    ' pointer to the object we are attached to
Private mstrKey                        As String  ' Key associatate to the form
Private m_hWnd                         As Long    ' handle for the form
Private m_Form                         As Form    ' pointer to the form object
Private mblnHasCloseButton             As Boolean
Private mstrVCtlKey                    As String  ' keep a weak reference to our clsControl object
Private m_SizingNow                    As Boolean
Private m_Align                        As eDAlignProperty
Private oSub                           As cSubclass
Private m_oIDF                         As IDoknForm
Private mblnMoveInProgress             As Boolean 'track undocked form move events
Implements iSubclass
'*****************************************************
Public Property Get Align() As eDAlignProperty
'*****************************************************
10   Align = m_Align
End Property
'*****************************************************
Public Property Let Align(New_Align As eDAlignProperty)
'*****************************************************
10   m_Align = New_Align
End Property
'*****************************************************
Public Function AllowDocking() As Boolean
'*****************************************************
   ' Purpose   : Returns true if the specified form is allowed to dock
   '** 11/30/06 Yorgi: Integrating DockedForms functionality
10   Select Case m_Align
        Case DAlignLeft
20         AllowDocking = m_Style And DSLeft
30      Case DAlignRight
40         AllowDocking = m_Style And DSRight
50      Case DAlignTop
60         AllowDocking = m_Style And DSTop
70      Case DAlignBottom
80         AllowDocking = m_Style And DSBottom
90      End Select
End Function
'*****************************************************
Public Property Get AttachToCtrlPtr() As Object
'*****************************************************
10   Set AttachToCtrlPtr = ObjectFromPtr(m_AttachToCtrlPtr)
End Property
'*****************************************************
Public Property Let AttachToCtrlPtr(oObj As Object)
'*****************************************************
10   m_AttachToCtrlPtr = ObjPtr(oObj)
End Property
'*****************************************************
Private Sub AttachToFormWnd()
'*****************************************************
   ' Purpose   : attach window messages to a specific hWnd
10   DetachFormWnd                                'detach any previous form within this host control?
20   If Not m_Form Is Nothing Then
30      m_hWnd = m_Form.hWnd
        'AppTrace m_constClassName, "AttachToFormWnd", "m_hWnd:" & CStr(m_hWnd)
        'Subclass the Form window
40      oSub.Subclass m_hWnd, Me
50      oSub.AddMsg m_hWnd, WM_SIZE, MSG_AFTER
60      oSub.AddMsg m_hWnd, WM_SYSCOMMAND, MSG_BEFORE
70      oSub.AddMsg m_hWnd, WM_DESTROY, MSG_BEFORE
        'now subclass the Host Container and track size changes
80      With m_ctlHostContainer
90         oSub.Subclass .hWnd, Me
100         oSub.AddMsg .hWnd, WM_SIZE, MSG_AFTER
110         End With
120      End If
End Sub
'*****************************************************
Private Sub Class_Initialize()
'*****************************************************
10   Const constSource As String = m_constClassName & ".Class_Initialize"
20   On Error GoTo Err_Class_Initialize
30   mblnStyleChanged = False
40   Set oSub = New cSubclass
50   Exit Sub
60   Err_Class_Initialize:
70   Err.Raise Err.Description, Err.Number, constSource:
End Sub
'*****************************************************
Private Sub Class_Terminate()
'*****************************************************
   ' Purpose   : Terminate the class (clean it up)
    'AppTrace m_constClassName, ".Class_Terminate", " begin"
10   DetachFormWnd                                ' DetachFormWnd all window messages
20   Set m_Form = Nothing
30   Set oSub = Nothing
40   Set m_ctlHostContainer = Nothing
     'AppTrace m_constClassName, ".Class_Terminate", " end"
End Sub
'*****************************************************
Friend Sub DetachFormWnd()
'*****************************************************
10   On Error Resume Next                         'just incase subclassing was already stopped
     'AppTrace m_constClassName, "DetachFormWnd", " m_hWnd:" & CStr(m_hWnd)
20   Set m_oIDF = Nothing
30   If (m_hWnd <> 0) Then
40      oSub.UnSubclass m_hWnd
50      m_hWnd = 0
60      oSub.UnSubclass m_ctlHostContainer.hWnd
70      End If
80   On Error GoTo 0
End Sub
'*****************************************************
Friend Property Get DockedForm() As Form
'*****************************************************
10   Set DockedForm = m_Form
End Property
'*****************************************************
Friend Property Let DockedForm(ByRef ofrm As Form)
'*****************************************************
10   Set m_Form = ofrm
20   AttachToFormWnd
End Property
'*****************************************************
Public Property Let FloatingExStyle(ByVal lFloatingExStyle As Long)
'*****************************************************
10   mlngFloatingExStyle = lFloatingExStyle
End Property
'*****************************************************
Public Property Get FloatingExStyle() As Long
'*****************************************************
10   FloatingExStyle = mlngFloatingExStyle
End Property
'*****************************************************
Public Property Get FloatingHeight() As Long
Attribute FloatingHeight.VB_Description = "Returns the floating height of the form"
'*****************************************************
10   FloatingHeight = mlngFloatingHeight
End Property
'*****************************************************
Public Property Let FloatingHeight(ByVal vData As Long)
'*****************************************************
10   mlngFloatingHeight = vData
End Property
'*****************************************************
Public Property Get FloatingLeft() As Long
Attribute FloatingLeft.VB_Description = "Returns the floating Left of the form"
'*****************************************************
10   FloatingLeft = mlngFloatingLeft
End Property
'*****************************************************
Public Property Let FloatingLeft(ByVal vData As Long)
'*****************************************************
10   mlngFloatingLeft = vData
End Property
'*****************************************************
Public Property Let FloatingStyle(ByVal lFloatingStyle As Long)
'*****************************************************
10   mlngFloatingStyle = lFloatingStyle
End Property
'*****************************************************
Public Property Get FloatingStyle() As Long
'*****************************************************
10   FloatingStyle = mlngFloatingStyle
End Property
'*****************************************************
Public Property Get FloatingTop() As Long
Attribute FloatingTop.VB_Description = "Returns the floating Top of the form"
'*****************************************************
10   FloatingTop = mlngFloatingTop
End Property
'*****************************************************
Public Property Let FloatingTop(ByVal vData As Long)
'*****************************************************
10   mlngFloatingTop = vData
End Property
'*****************************************************
Public Property Let FloatingWidth(ByVal vData As Long)
'*****************************************************
10   mlngFloatingWidth = vData
End Property
'*****************************************************
Public Property Get FloatingWidth() As Long
Attribute FloatingWidth.VB_Description = "Returns the floating Width of the form"
'*****************************************************
10   FloatingWidth = mlngFloatingWidth
End Property
'*****************************************************
Public Property Let FormParentHwnd(ByVal hWnd As Long)
'*****************************************************
10   mlngParenthWnd = hWnd
End Property
'*****************************************************
Public Property Get FormParentHwnd() As Long
'*****************************************************
10   FormParentHwnd = mlngParenthWnd
End Property
'*****************************************************
Public Property Let HasCloseButton(ByVal bHasCloseButton As Boolean)
'*****************************************************
10   mblnHasCloseButton = bHasCloseButton
End Property
'*****************************************************
Public Property Get HasCloseButton() As Boolean
'*****************************************************
10   HasCloseButton = mblnHasCloseButton
End Property
'*****************************************************
Friend Property Get HostContainer() As PictureBox
'*****************************************************
10   Set HostContainer = m_ctlHostContainer
End Property
'*****************************************************
Friend Property Set HostContainer(ctlHostContainer As PictureBox)
'*****************************************************
10   Set m_ctlHostContainer = ctlHostContainer
End Property
'*****************************************************
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns the form handle. This may not be used in api callings once the docked form is already subclassed and another subclassing of this form may crash Vb IDE or unpredictable results may take place."
'*****************************************************
10   hWnd = m_hWnd
End Property
'*****************************************************
Private Sub iSubclass_WndProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As eMsg, ByVal wParam As Long, ByVal lParam As Long)
'*****************************************************
   Dim retVal              As Long
   Dim uCmdType            As Long
10   On Error GoTo iSubclass_WndProc_Err
20   Select Case uMsg
        Case WM_SIZE
30         If lng_hWnd = m_ctlHostContainer.hWnd Then 'must be host container sizing
              'move form to fill up to host container
              'AppTrace m_constClassName, "iSubclass_WndProc", "WM_SIZE hWnd:" & CStr(lng_hWnd) & ", wParam:" & wParam & ", lHiParam:" & HiWord(lParam) & ", lLoParam:" & LoWord(lParam)
40            m_Form.Move 0, 0, m_ctlHostContainer.ScaleWidth, m_ctlHostContainer.ScaleHeight
50            End If
60      Case WM_MOVING
           'AppTrace m_constClassName, "iSubclass_WndProc", "WM_MOVING hWnd:" & CStr(lng_hWnd)
70         If Not mblnMoveInProgress Then
              'Notify ControlManager to begin the Move process
              'AppTrace m_constClassName, "iSubclass_WndProc", "WM_ENTERSIZEMOVE hWnd:" & CStr(lng_hWnd) & ", wParam:" & wParam & ", lParam:" & lParam
80            m_oIDF.WinEvent lng_hWnd, WM_ENTERSIZEMOVE, Me, wParam, lParam
90            mblnMoveInProgress = True
100            End If
            'Notify ControlManager we be Move'n
110         m_oIDF.WinEvent lng_hWnd, uMsg, Me, wParam, lParam
120      Case WM_EXITSIZEMOVE
            'AppTrace m_constClassName, "iSubclass_WndProc", "WM_EXITSIZEMOVE hWnd:" & CStr(lng_hWnd) & ", wParam:" & wParam & ", lParam:" & lParam
130         SaveFormFloatDimensions
140         If mblnMoveInProgress Then
               'Notify ControlManager to stand still for a while
150            mblnMoveInProgress = False
160            m_oIDF.WinEvent lng_hWnd, uMsg, Me, wParam, lParam
170            End If
180      Case WM_SYSCOMMAND
190         uCmdType = wParam And &HFFF0
200         If uCmdType = SC_MOVE Then
               'AppTrace m_constClassName, "iSubclass_WndProc", "WM_SYSCOMMAND SC_MOVE hWnd:" & CStr(lng_hWnd)
210            ReleaseCapture
220         ElseIf uCmdType = SC_CLOSE Then       'NOT UCMDTYPE...
               'AppTrace m_constClassName, "iSubclass_WndProc", "WM_SYSCOMMAND SC_CLOSE hWnd:" & CStr(lng_hWnd)
230            End If
240      Case WM_DESTROY
            'AppTrace m_constClassName, "iSubclass_WndProc", "WM_DESTROY hWnd:" & CStr(lng_hWnd)
            'user may have unloaded the form so we must detach now invalid references
250         m_oIDF.WinEvent lng_hWnd, uMsg, Me, wParam, lParam
260      End Select
270   iSubclass_WndProc_Exit:
280   On Error GoTo 0
290   Exit Sub
300   iSubclass_WndProc_Err:
310   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", iSubclass_WndProc", m_constClassName
320   Resume iSubclass_WndProc_Exit
End Sub
'*****************************************************
Public Property Get Key() As String
Attribute Key.VB_Description = "Returns the key of the form to be used with TabDock functions like FormShow() or FormHide. Instead of the key you may use the Form index to retrieve form reference"
'*****************************************************
10   Key = mstrKey
End Property
'*****************************************************
Public Property Let Key(vData As String)
'*****************************************************
10   mstrKey = vData
End Property
'*****************************************************
Public Property Let MinHeight(ByVal lngMinHeight As Long)
'*****************************************************
   ' Purpose    - Sets the minimum height of the form
   ' Input      - lngMinHeight(the new MinHeight propety value)
10   mlngMinHeight = lngMinHeight
End Property
'*****************************************************
Public Property Get MinHeight() As Long
'*****************************************************
   ' Purpose    - Returns the minimum height of the form
10   MinHeight = mlngMinHeight
End Property
'*****************************************************
Public Property Let MinWidth(ByVal lngMinWidth As Long)
'*****************************************************
   ' Purpose    - Sets the minimum width of the form
   ' Input      - lngMinWidth(the new MinWidth propety value)
10   mlngMinWidth = lngMinWidth
End Property
'*****************************************************
Public Property Get MinWidth() As Long
'*****************************************************
   ' Purpose    - Returns the minimum width of the form
10   MinWidth = mlngMinWidth
End Property
'*****************************************************
Public Property Get Object() As DokNForm
'*****************************************************
10   Set Object = Me
End Property
'*****************************************************
Public Property Set oIDF(objoIDF As IDoknForm)
'*****************************************************
10   Set m_oIDF = objoIDF
End Property
'*****************************************************
Public Property Get ParentForm() As Object
'*****************************************************
10   Set ParentForm = m_Form
End Property
'*****************************************************
Private Sub SaveFormFloatDimensions()
'*****************************************************
10   mlngFloatingWidth = m_Form.Width
20   mlngFloatingHeight = m_Form.Height
30   mlngFloatingLeft = m_Form.Left
40   mlngFloatingTop = m_Form.Top
End Sub
'*****************************************************
Public Property Let State(ByVal mNewState As eDockState)
'*****************************************************
10   If mNewState = DS_Docked Then
20      If m_State = DS_UnDocked Then             'make sure we were DS_UnDocked before
           'stop tracking moves while docked
30         oSub.DelMsg m_hWnd, WM_MOVING, MSG_BEFORE
40         oSub.DelMsg m_hWnd, WM_EXITSIZEMOVE, MSG_AFTER
50         End If
60   ElseIf mNewState = DS_UnDocked Then          'NOT MNEWSTATE...
        'start tracking moves to update float dimensions
70      oSub.AddMsg m_hWnd, WM_MOVING, MSG_BEFORE
80      oSub.AddMsg m_hWnd, WM_EXITSIZEMOVE, MSG_AFTER
90      End If
100   m_State = mNewState
End Property
'*****************************************************
Public Property Get State() As eDockState
Attribute State.VB_Description = "Returns the state of the form. Docked or Undocked."
'*****************************************************
10   State = m_State
End Property
'*****************************************************
Public Property Get Style() As eDockStyles
Attribute Style.VB_Description = "Returns or sets the form style. The style may allow or avoid some features of a docked window. Check all the possible styles to choose the apropriate one."
'*****************************************************
10   Style = m_Style
End Property
'*****************************************************
Public Property Let Style(ByVal vData As eDockStyles)
'*****************************************************
10   m_Style = vData
End Property
'*****************************************************
Public Property Get TBarPos() As TBarOrientation
'*****************************************************
10   TBarPos = m_TBarPos
End Property
'*****************************************************
Public Property Let TBarPos(vData As TBarOrientation)
'*****************************************************
10   m_TBarPos = vData
End Property
'*****************************************************
Public Property Let TBarType(ByVal eTBarType As TBarTypes)
'*****************************************************
10   m_TBarType = eTBarType
End Property
'*****************************************************
Public Property Get TBarType() As TBarTypes
'*****************************************************
10   TBarType = m_TBarType
End Property
'*****************************************************
Public Property Get VCtlKey() As String
'*****************************************************
10   VCtlKey = mstrVCtlKey
End Property
'*****************************************************
Public Property Let VCtlKey(sVCtlKey As String)
'*****************************************************
10   mstrVCtlKey = sVCtlKey
End Property
'*****************************************************
Friend Property Let Visible(vData As Boolean)
'*****************************************************
10   m_Form.Visible = vData
End Property
'*****************************************************
Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns wether the form is visible or not"
'*****************************************************
10   Visible = m_Form.Visible
End Property

' Yorgi's 4Matz [Feb 28,2007 23:58:49] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

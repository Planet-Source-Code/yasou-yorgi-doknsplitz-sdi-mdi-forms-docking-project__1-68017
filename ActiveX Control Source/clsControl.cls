VERSION 1.0 CLASS
BEGIN
MultiUse = -1                                     'True
Persistable = 0                                   'NotPersistable
DataBindingBehavior = 0                           'vbNone
DataSourceBehavior  = 0                           'vbNone
MTSTransactionMode  = 0                           'NotAnMTSObject
END
Attribute VB_Name = "clsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A class module to represent a ""virtual"" control"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************************
'** File Name     - clsControl.cls                                            **
'** Description   - A class module to represent a virtual control             **
'** Usage         - Instead of using a real control (like TextBox control)    **
'**                 properties directly, VB Control Manager control saves the **
'**                 necessary properties plus several custom properties in a  **
'**                 virtual control. Furthermore, VB Control Manager control  **
'**                 works with this virtual control to manipulate its         **
'**                 position and size and then write it back to the related   **
'                   real control.                                             **
'** Dependencies  - mdlGeneral                                                **
'** Public Members-                                                           **
'**   * Collections- -                                                        **
'**   * Objects    - -                                                        **
'**   * Properties - Bottom (r/o), Closed (r/o), Height (r/o), Id (def. r/o), **
'**                  IdSplBottom (r/o), IdSplLeft (r/o), IdSplRight (r/o),    **
'**                  IdSplTop (r/o), Left (r/o), MinHeight (r/o),             **
'**                  MinWidth (r/o), Name (r/o), Right (r/o),                 **
'**                  TitleBar_CloseVisible (r/w), TitleBar_Height (r/o),      **
'**                  TitleBar_Visible (r/w), Top (r/o), Width (r/o),          **
'**                  Xc (r/o), Yc (r/o)                                       **
'**   * Methods    - -                                                        **
'**   * Events     - -                                                        **
'** Friend Members-                                                           **
'**   * Collections- -                                                        **
'**   * Objects    - -                                                        **
'**   * Properties - Bottom (r/w), Closed (r/w), Height (r/w), Id (def. r/w), **
'**                  IdCtlFriendBottom (r/w), IdCtlFriendLeft (r/w),          **
'**                  IdCtlFriendRight (r/w), IdCtlFriendTop (r/w),            **
'**                  IdSplBottom (r/w), IdSplLeft (r/w), IdSplRight (r/w),    **
'**                  IdSplTop (r/w), Left (r/w), MinHeight (r/w),             **
'**                  MinWidth (r/w), Name (r/w), Parent (r/w), Right (r/w),   **
'**                  TitleBar_CloseVisible (r/w), TitleBar_Height (r/w),      **
'**                  TitleBar_Visible (r/w), TitleBar_VisibleHeight (r/o),    **
'**                  Top (r/w), Width (r/w), Xc (r/o), Yc (r/o)               **
'**   * Methods    - -                                                        **
'** Notes         - * Height and Yc properties value are generated from       **
'**                   Bottom and Top properties value                         **
'**                 * Width and Xc properties value are generated from Left   **
'**                   and Right properties value                              **
'** Last modified on September 2, 2003                                        **
'*******************************************************************************
Option Explicit
Private mlngptrCtl                     As Long    'weak reference the control
Private strDFKey                       As String  'DockedForm Key Index
Private mlngTbarIdx                    As Long    'TBar index
'--- Property Variables
Private mstrKey                        As String  'Control name
''''Private mlngIdx                        As Long    'my index in clsControls collection
Private mParent                        As clsControls
Private mblnClosed                     As Boolean
' Position and Size
Private mlngLeft                       As Long
Private mlngTop                        As Long
Private mlngRight                      As Long
Private mlngBottom                     As Long
' Minimum Size
Private mlngMinWidth                   As Long
Private mlngMinHeight                  As Long
'Title Bar
Private mblnTitleBar_CloseVisible      As Boolean
Private mlngTitleBar_Height            As Long
Private mblnTitleBar_Visible           As Boolean
Private mlngTitleBar_Type              As TBarTypes
Private mlngTitleBar_Pos               As TBarOrientation
' Friends Controls' Id
' Note- Friend controls are the nearest control around the control
Private msIdCtlFriendTop               As String
Private msIdCtlFriendRight             As String
Private msIdCtlFriendBottom            As String
Private msIdCtlFriendLeft              As String
' Friend Splitters' Id
' Note- Friend splitters (called splitters in this class) are the splitters
'       around the control
Private mIdSplTop                      As Long
Private mIdSplRight                    As Long
Private mIdSplBottom                   As Long
Private mIdSplLeft                     As Long
'--- Property Default Values
Private Const mconDefaultClosed        As Boolean = False
'--- Implements the clsTitleBar interface
Implements TitleBar
'*****************************************************
Public Property Get Bottom() As Long
Attribute Bottom.VB_Description = "Returns the distance between the internal bottom edge of the virtual control and the top edge of the related Control Manager object"
'*****************************************************
   ' Purpose    - Returns the distance between the internal bottom edge of the
   '              virtual control and the top edge of the related Control Manager
   '              object
10   Bottom = mlngBottom
End Property
'*****************************************************
Friend Property Let Bottom(ByVal lngBottom As Long)
'*****************************************************
   '------------------
   ' Class Properties
   '------------------
   ' Purpose    - Sets the distance between the internal bottom edge of the virtual
   '              control and the top edge of the related Control Manager object
   ' Input      - lngBottom (the new Bottom property value)
10   mlngBottom = lngBottom
End Property
'*****************************************************
Private Sub Class_Initialize()
'*****************************************************
10   mlngTbarIdx = gconUninitializedLong
20   strDFKey = vbNullString
30   mstrKey = vbNullString
40   mblnClosed = mconDefaultClosed
50   mlngMinWidth = 0
60   mlngMinHeight = 0
70   msIdCtlFriendTop = vbNullString
80   msIdCtlFriendRight = vbNullString
90   msIdCtlFriendBottom = vbNullString
100   msIdCtlFriendLeft = vbNullString
110   mIdSplTop = gconUninitializedLong
120   mIdSplRight = gconUninitializedLong
130   mIdSplBottom = gconUninitializedLong
140   mIdSplLeft = gconUninitializedLong
End Sub
'*****************************************************
Private Sub Class_Terminate()
'*****************************************************
10   Set mParent = Nothing
End Sub
'*****************************************************
Public Property Get Closed() As Boolean
Attribute Closed.VB_Description = "Returns a value that determines whether the control is closed"
'*****************************************************
   ' Purpose    - Returns a value that determines whether the virtual control is
   '              closed
10   Closed = mblnClosed
End Property
'*****************************************************
Friend Property Let Closed(ByVal blnClosed As Boolean)
'*****************************************************
   ' Purpose    - Sets a value that determines whether the virtual control is
   '              closed
   ' Input      - blnClosed (the new Closed property value)
10   mblnClosed = blnClosed
End Property
'*****************************************************
Public Property Get DFKey() As String
'*****************************************************
10   DFKey = strDFKey
End Property
'*****************************************************
Public Property Let DFKey(sDFKey As String)
'*****************************************************
10   strDFKey = sDFKey
End Property
'*****************************************************
Friend Property Let Height(ByVal lngHeight As Long)
'*****************************************************
   ' Purpose    - Sets the height of the virtual control
   ' Input      - lngHeight(the new Height property value)
10   mlngBottom = mlngTop + lngHeight
End Property
'*****************************************************
Public Property Get Height() As Long
Attribute Height.VB_Description = "Returns the height of the virtual control"
'*****************************************************
   ' Purpose    - Returns the height of the virtual control
10   Height = mlngBottom - mlngTop
End Property
'*****************************************************
Friend Property Let IdCtlFriendBottom(sIdCtlFriendBottom As String)
Attribute IdCtlFriendBottom.VB_Description = "Returns or sets the control's bottom-side friend control's Id"
'*****************************************************
   ' Purpose    - Sets the control's bottom-side friend control's Id
   ' Input      - lngIdCtlFriendBottom (the new IdCtlFriendBottom property value)
10   msIdCtlFriendBottom = sIdCtlFriendBottom
End Property
'*****************************************************
Friend Property Get IdCtlFriendBottom() As String
'*****************************************************
   ' Purpose    - Returns the control's bottom-side friend control's Id
10   IdCtlFriendBottom = msIdCtlFriendBottom
End Property
'*****************************************************
Friend Property Get IdCtlFriendLeft() As String
'*****************************************************
   ' Purpose    - Returns the control's left-side friend control's Id
10   IdCtlFriendLeft = msIdCtlFriendLeft
End Property
'*****************************************************
Friend Property Let IdCtlFriendLeft(sIdCtlFriendLeft As String)
Attribute IdCtlFriendLeft.VB_Description = "Returns or sets the control's left-side friend control's Id"
'*****************************************************
   ' Purpose    - Sets the control's left-side friend control's Id
   ' Input      - lngIdCtlFriendLeft (the new IdCtlFriendLeft property value)
10   msIdCtlFriendLeft = sIdCtlFriendLeft
End Property
'*****************************************************
Friend Property Get IdCtlFriendRight() As String
'*****************************************************
   ' Purpose    - Returns the control's right-side friend control's Id
10   IdCtlFriendRight = msIdCtlFriendRight
End Property
'*****************************************************
Friend Property Let IdCtlFriendRight(sIdCtlFriendRight As String)
Attribute IdCtlFriendRight.VB_Description = "Returns or sets the control's right-side friend control's Id"
'*****************************************************
   ' Purpose    - Sets the control's right-side friend control's Id
   ' Input      - lngIdCtlFriendRight (the new IdCtlFriendRight property value)
10   msIdCtlFriendRight = sIdCtlFriendRight
End Property
'*****************************************************
Friend Property Get IdCtlFriendTop() As String
'*****************************************************
   ' Purpose    - Returns the control's top-side friend control's Id
10   IdCtlFriendTop = msIdCtlFriendTop
End Property
'*****************************************************
Friend Property Let IdCtlFriendTop(sIdCtlFriendTop As String)
Attribute IdCtlFriendTop.VB_Description = "Returns or sets the control's top-side friend control's Id"
'*****************************************************
   ' Purpose    - Sets the control's top-side friend control's Id
   ' Input      - lngIdCtlFriendTop (the new IdCtlFriendTop property value)
10   msIdCtlFriendTop = sIdCtlFriendTop
End Property
'*****************************************************
Friend Property Let IdSplBottom(ByVal lngIdSplBottom As Long)
'*****************************************************
   ' Purpose    - Sets a value that uniquely identifies the virtual splitter which
   '              its movement will change the value of the virtual control's
   '              Bottom property
   ' Input      - lngIdSplFriendBottom (the new IdSplFriendBottom property value)
10   mIdSplBottom = lngIdSplBottom
End Property
'*****************************************************
Public Property Get IdSplBottom() As Long
Attribute IdSplBottom.VB_Description = "Returns a value that uniquely identifies the virtual splitter which its movement will change the value of the virtual control's Bottom property"
'*****************************************************
   ' Purpose    - Returns a value that uniquely identifies the virtual splitter
   '              which its movement will change the value of the virtual control's
   '              Bottom property
10   IdSplBottom = mIdSplBottom
End Property
'*****************************************************
Friend Property Let IdSplLeft(ByVal lngIdSplLeft As Long)
'*****************************************************
   ' Purpose    - Sets a value that uniquely identifies the virtual splitter which
   '              its movement will change the value of the virtual control's Left
   '              property
   ' Input      - lngIdSplFriendLeft (the new IdSplFriendLeft property value)
10   mIdSplLeft = lngIdSplLeft
End Property
'*****************************************************
Public Property Get IdSplLeft() As Long
Attribute IdSplLeft.VB_Description = "Returns a value that uniquely identifies the virtual splitter which its movement will change the value of the virtual control's Left property"
'*****************************************************
   ' Purpose    - Returns a value that uniquely identifies the virtual splitter
   '              which its movement will change the value of the virtual control's
   '              Left property
10   IdSplLeft = mIdSplLeft
End Property
'*****************************************************
Public Property Get IdSplRight() As Long
Attribute IdSplRight.VB_Description = "Returns a value that uniquely identifies the virtual splitter which its movement will change the value of the virtual control's Right property"
'*****************************************************
   ' Purpose    - Returns a value that uniquely identifies the virtual splitter
   '              which its movement will change the value of the virtual control's
   '              Right property
10   IdSplRight = mIdSplRight
End Property
'*****************************************************
Friend Property Let IdSplRight(ByVal lngIdSplRight As Long)
'*****************************************************
   ' Purpose    - Sets a value that uniquely identifies the virtual splitter which
   '              its movement will change the value of the virtual control's Right
   '              property
   ' Input      - lngIdSplFriendRight (the new IdSplFriendRight property value)
10   mIdSplRight = lngIdSplRight
End Property
'*****************************************************
Public Property Get IdSplTop() As Long
Attribute IdSplTop.VB_Description = "Returns a value that uniquely identifies the virtual splitter which its movement will change the value of the virtual control's Top property"
'*****************************************************
   ' Purpose    - Returns a value that uniquely identifies the virtual splitter
   '              which its movement will change the value of the virtual control's
   '              Top property
10   IdSplTop = mIdSplTop
End Property
'*****************************************************
Friend Property Let IdSplTop(ByVal lngIdSplTop As Long)
'*****************************************************
   ' Purpose    - Sets a value that uniquely identifies the virtual splitter which
   '              its movement will change the value of the virtual control's Top
   '              property
   ' Input      - lngIdSplFriendTop (the new IdSplFriendTop property value)
10   mIdSplTop = lngIdSplTop
End Property
'*****************************************************
Friend Property Let Key(sId As String)
'*****************************************************
   ' Purpose    - Sets a value that uniquely identifies the virtual control
   ' Input      - lngId(the new Id property value)
10   mstrKey = sId
End Property
'*****************************************************
Public Property Get Key() As String '4Matz:Changed
Attribute Key.VB_UserMemId = 0
Attribute Key.VB_Description = "Returns a value that uniquely identifies the virtual control"
'*****************************************************
   ' Purpose    - Returns a value that uniquely identifies the virtual control
10   Key = mstrKey
End Property
'*****************************************************
Public Property Get Left() As Long
Attribute Left.VB_Description = "Returns the distance between the internal left edge of the virtual control and the left edge of the related Control Manager object"
'*****************************************************
   ' Purpose    - Returns the distance between the internal left edge of the
   '              virtual control and the left edge of the related Control Manager
   '              object
10   Left = mlngLeft
End Property
'*****************************************************
Friend Property Let Left(ByVal lngLeft As Long)
'*****************************************************
   ' Purpose    - Sets the distance between the internal left edge of the virtual
   '              control and the left edge of the related Control Manager object
   ' Input      - lngLeft (the new Left property value)
10   mlngLeft = lngLeft
End Property
'*****************************************************
Friend Property Let MinHeight(ByVal lngMinHeight As Long)
'*****************************************************
   ' Purpose    - Sets the minimum height of the virtual control
   ' Input      - lngMinHeight(the new MinHeight propety value)
10   mlngMinHeight = lngMinHeight
End Property
'*****************************************************
Public Property Get MinHeight() As Long
Attribute MinHeight.VB_Description = "Returns the minimum height of the virtual control"
'*****************************************************
   ' Purpose    - Returns the minimum height of the virtual control
10   MinHeight = mlngMinHeight + mlngTitleBar_Height
End Property
'*****************************************************
Friend Property Let MinWidth(ByVal lngMinWidth As Long)
'*****************************************************
   ' Purpose    - Sets the minimum width of the virtual control
   ' Input      - lngMinWidth(the new MinWidth propety value)
10   mlngMinWidth = lngMinWidth
End Property
'*****************************************************
Public Property Get MinWidth() As Long
Attribute MinWidth.VB_Description = "Returns the minimum width of the virtual control"
'*****************************************************
   ' Purpose    - Returns the minimum width of the control
10   MinWidth = mlngMinWidth
End Property
'*****************************************************
Friend Property Set Parent(ByVal ctlsParent As clsControls)
'*****************************************************
   ' Purpose    - Sets the collection that contain this class
   ' Input      - splsParent (the new Parent property value)
10   Set mParent = ctlsParent
End Property
'*****************************************************
Friend Property Get Parent() As clsControls
Attribute Parent.VB_Description = "Returns the collection that contain this class"
'*****************************************************
   ' Purpose    - Returns the collection that contain this class
10   Set Parent = mParent
End Property
'*****************************************************
Friend Property Let refCtlObj(llpCCObj As Control)  '4Matz:New
'*****************************************************
10   mlngptrCtl = ObjPtr(llpCCObj)
End Property
'*****************************************************
Friend Property Get refCtlObj() As Control '4Matz:New
'*****************************************************
10   Set refCtlObj = ObjectFromPtr(mlngptrCtl)
End Property
'*****************************************************
Public Property Get Right() As Long
Attribute Right.VB_Description = "Returns the distance between the internal right edge of the virtual control and the left edge of the related Control Manager object"
'*****************************************************
   ' Purpose    - Returns the distance between the internal right edge of the
   '              virtual control and the left edge of the related Control Manager
   '              object
10   Right = mlngRight
End Property
'*****************************************************
Friend Property Let Right(ByVal lngRight As Long)
'*****************************************************
   ' Purpose    - Sets the distance between the internal right edge of the virtual
   '              control and the left edge of the related Control Manager object
   ' Input      - lngRight (the new Right property value)
10   mlngRight = lngRight
End Property
'*****************************************************
Public Property Let TbarIdx(ByVal lTbarIdx As Long)
'*****************************************************
10   mlngTbarIdx = lTbarIdx
End Property
'*****************************************************
Public Property Get TbarIdx() As Long
'*****************************************************
10   TbarIdx = mlngTbarIdx
End Property
'*****************************************************
Public Property Let TitleBar_CloseVisible(ByVal blnTitleBar_CloseVisible As Boolean)
Attribute TitleBar_CloseVisible.VB_Description = "Returns/sets a value that determines whether a close button in the virtual control title bar is visible"
'*****************************************************
   ' Purpose    - Sets a value that determines whether a close button in the
   '              virtual control title bar is visible
   ' Input      - blnTitleBar_CloseVisible (the new TitleBar_CloseVisible property
   '                                        value)
10   mblnTitleBar_CloseVisible = blnTitleBar_CloseVisible
20   Parent.TitleBar_CloseVisible(mstrKey) = mblnTitleBar_CloseVisible
End Property
'*****************************************************
Public Property Get TitleBar_CloseVisible() As Boolean
'*****************************************************
   ' Purpose    - Returns a value that determines whether a close button in the
   '              virtual control title bar is visible
10   TitleBar_CloseVisible = mblnTitleBar_CloseVisible
End Property
'*****************************************************
Friend Property Let TitleBar_Height(ByVal lngTitleBar_Height As Long)
'*****************************************************
   ' Purpose    - Sets the height of the virtual control title bar
   ' Input      - lngTitleBar_Height (the new TitleBar_Height property value)
10   mlngTitleBar_Height = lngTitleBar_Height
End Property
'*****************************************************
Public Property Get TitleBar_Height() As Long
'*****************************************************
   ' Purpose    - Returns the height of the visible part of the virtual control
   '              title bar
10   If mblnTitleBar_Visible Then
20      TitleBar_Height = mlngTitleBar_Height
30      End If
End Property
'*****************************************************
Public Property Get TitleBar_Position() As TBarOrientation
'*****************************************************
10   TitleBar_Position = mlngTitleBar_Pos
End Property
'*****************************************************
Public Property Let TitleBar_Position(ByVal lTBarPos As TBarOrientation)
'*****************************************************
10   mlngTitleBar_Pos = lTBarPos
End Property
'*****************************************************
Public Property Get TitleBar_TBarType() As TBarTypes
'*****************************************************
10   TitleBar_TBarType = mlngTitleBar_Type
End Property
'*****************************************************
Public Property Let TitleBar_TBarType(ByVal lTBarType As TBarTypes)
'*****************************************************
   ' Purpose    - Sets the TBarType of the virtual control title bar
   ' Input      - lTBarType (see mdlGeneral.TBarTypes)
10   mlngTitleBar_Type = lTBarType
End Property
'*****************************************************
Public Property Let TitleBar_Visible(ByVal blnTitleBar_Visible As Boolean)
'*****************************************************
   ' Purpose    - Sets a value that determines whether the virtual control title bar is visible
   ' Input      - blnblnTitleBar_Visible (the new blnTitleBar_Visible property value)
10   If mblnTitleBar_Visible <> blnTitleBar_Visible Then
20      mblnTitleBar_Visible = blnTitleBar_Visible
30      If Parent.IsValid Then
40         Parent.TitleBar_Visible(mstrKey) = mblnTitleBar_Visible
50      Else
60         mblnTitleBar_Visible = Not mblnTitleBar_Visible
70         End If
80      End If
End Property
'*****************************************************
Public Property Get TitleBar_Visible() As Boolean
'*****************************************************
   ' Purpose    - Returns a value that determines whether the virtual control title bar is visible
10   TitleBar_Visible = mblnTitleBar_Visible
End Property
'*****************************************************
Friend Property Get TitleBar_VisibleHeight() As Long
'*****************************************************
   ' Purpose    - Returns the height of the title bar
10   TitleBar_VisibleHeight = mlngTitleBar_Height
End Property
'*****************************************************
Friend Property Let Top(ByVal lngTop As Long)
'*****************************************************
   ' Purpose    - Sets the distance between the internal top edge of the virtual
   '              control and the top edge of the related Control Manager object
   ' Input      - lngTop (the new Top property value)
10   mlngTop = lngTop
End Property
'*****************************************************
Public Property Get Top() As Long
Attribute Top.VB_Description = "Returns the distance between the internal top edge of the virtual control and the top edge of the related Control Manager object"
'*****************************************************
   ' Purpose    - Returns the distance between the internal top edge of the virtual
   '              control and the top edge of the related Control Manager object
10   Top = mlngTop
End Property
'*****************************************************
Public Property Get Width() As Long
Attribute Width.VB_Description = "Returns the width of the virtual control"
'*****************************************************
   ' Purpose    - Returns the width of the virtual control
10   Width = mlngRight - mlngLeft
End Property
'*****************************************************
Friend Property Let Width(ByVal lngWidth As Long)
'*****************************************************
   ' Purpose    - Sets the width of the virtual control
   ' Input      - lngWidth(the new Width property value)
10   mlngRight = mlngLeft + lngWidth
End Property
'*****************************************************
Public Property Get Xc() As Long
Attribute Xc.VB_Description = "Returns the x-coordinate of the virtual control center"
'*****************************************************
   ' Purpose    - Returns the x-coordinate of the virtual control center
10   Xc = CLng((mlngLeft + mlngRight) / 2)
End Property
'*****************************************************
Public Property Get Yc() As Long
Attribute Yc.VB_Description = "Returns the y-coordinate of the virtual control center"
'*****************************************************
   ' Purpose    - Returns the y-coordinate of the virtual control center
10   Yc = CLng((mlngTop + mlngBottom) / 2)
End Property

' Yorgi's 4Matz [Feb 28,2007 23:58:48] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

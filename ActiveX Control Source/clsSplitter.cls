VERSION 1.0 CLASS
BEGIN
MultiUse = -1                                     'True
Persistable = 0                                   'NotPersistable
DataBindingBehavior = 0                           'vbNone
DataSourceBehavior  = 0                           'vbNone
MTSTransactionMode  = 0                           'NotAnMTSObject
END
Attribute VB_Name = "clsSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A class module to represent a ""virtual"" splitter"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************************
'** File Name     : clsSplitter.cls                                           **
'** Description   : A class module to represent a virtual splitter            **
'** Usage         : VB Control Manager control uses PictureBox control as a   **
'**                 splitter. Instead of using PictureBox properties directly **
'**                 VB Control Manager control saves the necessary properties **
'**                 plus several custom properties in a virtual splitter.     **
'**                 Furthermore, VB Control Manager control works with this   **
'**                 virtual splitter to manipulate its position and then      **
'**                 write it back to the PictureBox control.                  **
'** Dependencies  : clsIds, clsId, mdlGeneral                                 **
'** Public Members:                                                           **
'**   * Collections: IdsCtlBottom (r/o), IdsCtlLeft (r/o), IdsCtlRight (r/o), **
'**                  IdsCtlTop (r/o), IdsSplBottom (r/o), IdsSplLeft (r/o),   **
'**                  IdsSplRight (r/o), IdsSplTop (r/o)                       **
'**   * Object     : -                                                        **
'**   * Properties : ActiveColor (r/w), BackColor (r/w), Bottom (r/o),        **
'**                  ClipCursor (r/w), Enable (r/w), Height (r/o),            **
'**                  Id (def. r/o), Left (r/o), LiveUpdate (r/w),             **
'**                  MaxXc (r/o), MaxYc (r/o), MinXc (r/o), MinYc (r/o),      **
'**                  Orientation (r/o), Right (r/o), Top (r/o), Width (r/o),  **
'**                  Xc (r/o), Yc (r/o)                                       **
'**   * Methods    : -                                                        **
'**   * Events     : -                                                        **
'** Friend Members:                                                           **
'**   * Collections: IdsCtlBottom (r/o), IdsCtlLeft (r/o), IdsCtlRight (r/o), **
'**                  IdsCtlTop (r/o), IdsSplBottom (r/o), IdsSplLeft (r/o),   **
'**                  IdsSplRight (r/o), IdsSplTop (r/o), Parent (r/o)         **
'**   * Object     : -                                                        **
'**   * Properties : ActiveColor (r/w), BackColor (r/w), Bottom (r/w),        **
'**                  ClipCursor (r/w), Enable (r/w), Height (r/w),            **
'**                  Id (def. r/w), IdCtlFriendBottom (r/w),                  **
'**                  IdCtlFriendLeft (r/w), IdCtlFriendRight (r/w),           **
'**                  IdCtlFriendTop (r/w), Left (r/w), LiveUpdate (r/w),      **
'**                  MaxXc (r/w), MaxYc (r/w), MinXc (r/w),  MinYc (r/w),     **
'**                  Orientation (r/w), Parent (r/w), Right (r/w), Top (r/w), **
'**                  Width (r/w), Xc (r/w), Yc (r/w)                          **
'**   * Methods    : -                                                        **
'** Notes         : * Left and Right properties value are generated from Xc   **
'**                   and Width properties value                              **
'**                 * Top and Bottom properties value are generated from Yc   **
'**                   and Height properties value                             **
'** Last modified on September 2, 2003                                        **
'*******************************************************************************
Option Explicit
'--- Public Type Declaration
Public Enum genmSplitterOrientation
   orHorizontal
   orVertical
End Enum
'--- Property Variables
Private mId                            As Long
Private mParent                        As clsSplitters
Private mudeOrientation                As genmSplitterOrientation
' Property variables which appears in the property page
Private mblnClipCursor                 As Boolean
Private mblnEnable                     As Boolean
Private mblnLiveUpdate                 As Boolean
Private mlngActiveColor                As OLE_COLOR
Private mlngBackColor                  As OLE_COLOR
' Position and Size
Private mlngXc                         As Long
Private mlngYc                         As Long
Private mlngWidth                      As Long
Private mlngHeight                     As Long
' Splitter's moving region
Private mlngMinXc                      As Long
Private mlngMinYc                      As Long
Private mlngMaxXc                      As Long
Private mlngMaxYc                      As Long
' The controls which is effected by the splitter's motion
Private mIdsCtlTop                     As clsIds
Private mIdsCtlRight                   As clsIds
Private mIdsCtlBottom                  As clsIds
Private mIdsCtlLeft                    As clsIds
' The splitters which is effected by the splitter's motion
Private mIdsSplTop                     As clsIds
Private mIdsSplRight                   As clsIds
Private mIdsSplBottom                  As clsIds
Private mIdsSplLeft                    As clsIds
' Friends Controls' Id
' Note: Friend controls are the nearest control around the splitter
Private mIdCtlFriendTop                As Long
Private mIdCtlFriendRight              As Long
Private mIdCtlFriendBottom             As Long
Private mIdCtlFriendLeft               As Long
'----------------------------------
' Class Constructor and Destructor
'----------------------------------
'*****************************************************
Public Property Let ActiveColor(ByVal lngActiveColor As OLE_COLOR)
Attribute ActiveColor.VB_Description = "Returns/sets the background color used to display the splitters when the user moves it in none live update mode"
'*****************************************************
   '------------------
   ' Class Properties
   '------------------
   ' Purpose    - Sets the background color used to display the splitter when the
   '              user moves it in none live update mode
   ' Input      - lngActiveColor (the new ActiveColor property value)
10   mlngActiveColor = lngActiveColor
End Property
'*****************************************************
Public Property Get ActiveColor() As OLE_COLOR
'*****************************************************
   ' Purpose    - Returns the background color used to display the splitters when
   '              the user moves it in none live update mode
10   ActiveColor = mlngActiveColor
End Property
'*****************************************************
Public Property Let BackColor(lngBackColor As OLE_COLOR)
Attribute BackColor.VB_Description = "Sets/returns the background color used to display the splitter"
'*****************************************************
   ' Purpose    - Sets the background color used to display the splitter
   ' Input      - lngBackColor (the new BackColor property value)
10   mlngBackColor = lngBackColor
20   Parent.BackColor(mId) = mlngBackColor
End Property
'*****************************************************
Public Property Get BackColor() As OLE_COLOR
'*****************************************************
   ' Purpose    - Returns the background color used to display the splitters
10   BackColor = mlngBackColor
End Property
'*****************************************************
Friend Property Let Bottom(ByVal lngBottom As Long)
'*****************************************************
   ' Purpose    - Sets the distance between the internal bottom edge of the virtual
   '              splitter and the top edge of the related Control Manager object
   ' Input      - lngBottom (the new Bottom property value)
   Dim lngTop              As Long
10   lngTop = Top
20   mlngHeight = lngBottom - lngTop
30   mlngYc = CLng((lngTop + lngBottom) / 2)
End Property
'*****************************************************
Public Property Get Bottom() As Long
Attribute Bottom.VB_Description = "Returns the distance between the internal bottom edge of the virtual splitter and the top edge of the related Control Manager object"
'*****************************************************
   ' Purpose    - Returns the distance between the internal bottom edge of the
   '              virtual splitter and the top edge of the related Control Manager
   '              object
10   Bottom = CLng(mlngYc + (mlngHeight / 2))
End Property
'*****************************************************
Private Sub Class_Initialize()
'*****************************************************
10   mId = gconUninitializedLong
20   Set mIdsCtlTop = New clsIds
30   Set mIdsCtlRight = New clsIds
40   Set mIdsCtlBottom = New clsIds
50   Set mIdsCtlLeft = New clsIds
60   Set mIdsSplTop = New clsIds
70   Set mIdsSplRight = New clsIds
80   Set mIdsSplBottom = New clsIds
90   Set mIdsSplLeft = New clsIds
100   mIdCtlFriendTop = gconUninitializedLong
110   mIdCtlFriendRight = gconUninitializedLong
120   mIdCtlFriendBottom = gconUninitializedLong
130   mIdCtlFriendLeft = gconUninitializedLong
      ' Initialize the MinXc and MinYc property to the minimum value and the MaxXc
      '   and MaxYc tot the maximum value to make sure that these values will be
      '   replaced with the real minimum and maximum values
140   mlngMinXc = 0
150   mlngMinYc = 0
160   mlngMaxXc = gconLngInfinite
170   mlngMaxYc = gconLngInfinite
End Sub
'*****************************************************
Private Sub Class_Terminate()
'*****************************************************
10   Set mIdsCtlTop = Nothing
20   Set mIdsCtlRight = Nothing
30   Set mIdsCtlBottom = Nothing
40   Set mIdsCtlLeft = Nothing
50   Set mIdsSplTop = Nothing
60   Set mIdsSplRight = Nothing
70   Set mIdsSplBottom = Nothing
80   Set mIdsSplLeft = Nothing
End Sub
'*****************************************************
Public Property Let ClipCursor(ByVal blnClipCursor As Boolean)
Attribute ClipCursor.VB_Description = "Returns/sets a value that determines whether the mouse pointer is confined to the virtual splitter minimum and maximum x- and y-coordinate when the user moves the splitter"
'*****************************************************
   ' Purpose    - Sets a value that determines whether the mouse pointer is confined
   '              to the virtual splitter minimum and maximum x- and y-coordinate
   '              when the user moves the splitter
   ' Input      - blnClipCursor (the new ClipCursor property value)
10   mblnClipCursor = blnClipCursor
End Property
'*****************************************************
Public Property Get ClipCursor() As Boolean
'*****************************************************
   ' Purpose    - Returns/sets a value that determines whether the mouse pointer is
   '              confined to the virtual splitter minimum and maximum x- and
   '              y-coordinate when the user moves the splitter
10   ClipCursor = mblnClipCursor
End Property
'*****************************************************
Public Property Let Enable(ByVal blnEnable As Boolean)
Attribute Enable.VB_Description = "Returns/sets a value that determines whether the splitter is movable"
'*****************************************************
   ' Purpose    - Sets a value that determines whether the virtual splitter is
   '              movable
   ' Input      - blnEnable (the new Enable property value)
10   mblnEnable = blnEnable
20   Parent.Enable(mId) = mblnEnable
End Property
'*****************************************************
Public Property Get Enable() As Boolean
'*****************************************************
   ' Purpose    - Returns a value that determines whether the virtual splitter is
   '              movable
10   Enable = mblnEnable
End Property
'*****************************************************
Friend Property Let Height(ByVal lngHeight As Long)
'*****************************************************
   ' Purpose    - Sets the height of the virtual splitter
   ' Input      - lngHeight(the new Height property value)
10   mlngHeight = lngHeight
End Property
'*****************************************************
Public Property Get Height() As Long
Attribute Height.VB_Description = "Returns the height of the splitter"
'*****************************************************
   ' Purpose    - Returns the height of the virtual splitter
10   Height = mlngHeight
End Property
'*****************************************************
Friend Property Let Id(ByVal lngId As Long)
'*****************************************************
   ' Purpose    - Sets the id of the splitter
   ' Input      - lngId (the new Id property value)
10   mId = lngId
End Property
'*****************************************************
Public Property Get Id() As Long
Attribute Id.VB_UserMemId = 0
Attribute Id.VB_Description = "Returns a value that uniquely identifies the virtual splitter"
'*****************************************************
   ' Purpose    - Returns the id of the splitter
10   Id = mId
End Property
'*****************************************************
Friend Property Let IdCtlFriendBottom(ByVal lngIdCtlFriendBottom As Long)
Attribute IdCtlFriendBottom.VB_Description = "Returns or sets the id of the splitter's nearest bottom-side control"
'*****************************************************
   ' Purpose    - Sets the id of the splitter's nearest bottom-side control
   ' Input      - lngIdCtlFriendBottom (the new IdCtlFriendBottom property values)
10   mIdCtlFriendBottom = lngIdCtlFriendBottom
End Property
'*****************************************************
Friend Property Get IdCtlFriendBottom() As Long
'*****************************************************
   ' Purpose    - Returns the id of the splitter's nearest bottom-side control
10   IdCtlFriendBottom = mIdCtlFriendBottom
End Property
'*****************************************************
Friend Property Let IdCtlFriendLeft(ByVal lngIdCtlFriendLeft As Long)
Attribute IdCtlFriendLeft.VB_Description = "Returns or sets the id of the splitter's nearest left-side control"
'*****************************************************
   ' Purpose    - Sets the id of the splitter's nearest left-side control
   ' Input      - lngIdCtlFriendLeft (the new IdCtlFriendLeft property values)
10   mIdCtlFriendLeft = lngIdCtlFriendLeft
End Property
'*****************************************************
Friend Property Get IdCtlFriendLeft() As Long
'*****************************************************
   ' Purpose    - Returns the id of the splitter's nearest left-side control
10   IdCtlFriendLeft = mIdCtlFriendLeft
End Property
'*****************************************************
Friend Property Let IdCtlFriendRight(ByVal lngIdCtlFriendRight As Long)
Attribute IdCtlFriendRight.VB_Description = "Returns or sets the id of the splitter's nearest right-side control"
'*****************************************************
   ' Purpose    - Sets the id of the splitter's nearest right-side control
   ' Input      - lngIdCtlFriendRight (the new IdCtlFriendRight property values)
10   mIdCtlFriendRight = lngIdCtlFriendRight
End Property
'*****************************************************
Friend Property Get IdCtlFriendRight() As Long
'*****************************************************
   ' Purpose    - Returns the id of the splitter's nearest right-side control
10   IdCtlFriendRight = mIdCtlFriendRight
End Property
'*****************************************************
Friend Property Let IdCtlFriendTop(ByVal lngIdCtlFriendTop As Long)
Attribute IdCtlFriendTop.VB_Description = "Returns or sets the id of the splitter's nearest top-side control"
'*****************************************************
   ' Purpose    - Sets the id of the splitter's nearest top-side control
   ' Input      - lngIdCtlFriendTop (the new IdCtlFriendTop property values)
10   mIdCtlFriendTop = lngIdCtlFriendTop
End Property
'*****************************************************
Friend Property Get IdCtlFriendTop() As Long
'*****************************************************
   ' Purpose    - Returns the id of the splitter's nearest top-side control
10   IdCtlFriendTop = mIdCtlFriendTop
End Property
'*****************************************************
Friend Property Set IdsCtlBottom(oidsCtlBottom As clsIds)
'*****************************************************
   ' Purpose    - Sets the Id collection of all virtual controls which are effected
   '              by the virtual splitter's down-movement
   ' Input      - oidsCtlBottom (the new IdsCtlBottom property values)
10   Set mIdsCtlBottom = oidsCtlBottom
End Property
'*****************************************************
Public Property Get IdsCtlBottom() As clsIds
Attribute IdsCtlBottom.VB_Description = "Returns the Id collection of all virtual controls which are effected by the virtual splitter's down-movement"
'*****************************************************
   ' Purpose    - Returns the Id collection of all virtual controls which are
   '              effected by the virtual splitter's down-movement
10   Set IdsCtlBottom = mIdsCtlBottom
End Property
'*****************************************************
Friend Property Set IdsCtlLeft(oidsCtlLeft As clsIds)
'*****************************************************
   ' Purpose    - Sets the Id collection of all virtual controls which are effected
   '              by the virtual splitter's left-movement
   ' Input      - oidsCtlLeft (the new IdsCtlLeft property values)
10   Set mIdsCtlLeft = oidsCtlLeft
End Property
'*****************************************************
Public Property Get IdsCtlLeft() As clsIds
Attribute IdsCtlLeft.VB_Description = "Returns the Id collection of all virtual controls which are effected by the virtual splitter's left-movement"
'*****************************************************
   ' Purpose    - Returns the Id collection of all virtual controls which are
   '              effected by the virtual splitter's left-movement
10   Set IdsCtlLeft = mIdsCtlLeft
End Property
'*****************************************************
Friend Property Set IdsCtlRight(oidsCtlRight As clsIds)
'*****************************************************
   ' Purpose    - Sets the Id collection of all virtual controls which are effected
   '              by the virtual splitter's right-movement
   ' Input      - oidsCtlRight (the new IdsCtlRight property values)
10   Set mIdsCtlRight = oidsCtlRight
End Property
'*****************************************************
Public Property Get IdsCtlRight() As clsIds
Attribute IdsCtlRight.VB_Description = "Returns the Id collection of all virtual controls which are effected by the virtual splitter's right-movement"
'*****************************************************
   ' Purpose    - Returns the Id collection of all virtual controls which are
   '              effected by the virtual splitter's right-movement
10   Set IdsCtlRight = mIdsCtlRight
End Property
'*****************************************************
Friend Property Set IdsCtlTop(oidsCtlTop As clsIds)
'*****************************************************
   ' Purpose    - Sets the Id collection of all virtual controls which are effected
   '              by the virtual splitter's up-movement
   ' Input      - oidsCtlTop (the new IdsCtlTop property values)
10   Set mIdsCtlTop = oidsCtlTop
End Property
'*****************************************************
Public Property Get IdsCtlTop() As clsIds
Attribute IdsCtlTop.VB_Description = "Returns the Id collection of all virtual controls which are effected by the virtual splitter's up-movement"
'*****************************************************
   ' Purpose    - Returns the Id collection of all virtual controls which are
   '              effected by the virtual splitter's up-movement
10   Set IdsCtlTop = mIdsCtlTop
End Property
'*****************************************************
Friend Property Set IdsSplBottom(oidsSplBottom As clsIds)
'*****************************************************
   ' Purpose    - Sets the Id collection of all virtual splitters which are
   '              effected by the virtual splitter's down-movement
   ' Input      - oidsSplBottom (the new IdsSplBottom property values)
10   Set mIdsSplBottom = oidsSplBottom
End Property
'*****************************************************
Public Property Get IdsSplBottom() As clsIds
Attribute IdsSplBottom.VB_Description = "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's down-movement"
'*****************************************************
   ' Purpose    - Returns the Id collection of all virtual splitters which are
   '              effected by the virtual splitter's down-movement
10   Set IdsSplBottom = mIdsSplBottom
End Property
'*****************************************************
Friend Property Set IdsSplLeft(oidsSplLeft As clsIds)
'*****************************************************
   ' Purpose    - Sets the Id collection of all virtual splitters which are
   '              effected by the virtual splitter's left-movement
   ' Input      - oidsSplLeft (the new IdsSplLeft property values)
10   Set mIdsSplLeft = oidsSplLeft
End Property
'*****************************************************
Public Property Get IdsSplLeft() As clsIds
Attribute IdsSplLeft.VB_Description = "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's left-movement"
'*****************************************************
   ' Purpose    - Returns the Id collection of all virtual splitters which are
   '              effected by the virtual splitter's left-movement
10   Set IdsSplLeft = mIdsSplLeft
End Property
'*****************************************************
Friend Property Set IdsSplRight(oidsSplRight As clsIds)
'*****************************************************
   ' Purpose    - Sets the Id collection of all virtual splitters which are
   '              effected by the virtual splitter's right-movement
   ' Input      - oidsSplRight (the new IdsSplRight property values)
10   Set mIdsSplRight = oidsSplRight
End Property
'*****************************************************
Public Property Get IdsSplRight() As clsIds
Attribute IdsSplRight.VB_Description = "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's right-movement"
'*****************************************************
   ' Purpose    - Returns the Id collection of all virtual splitters which are
   '              effected by the virtual splitter's right-movement
10   Set IdsSplRight = mIdsSplRight
End Property
'*****************************************************
Friend Property Set IdsSplTop(oidsSplTop As clsIds)
'*****************************************************
   ' Purpose    - Sets the Id collection of all virtual splitters which are
   '              effected by the virtual splitter's up-movement
   ' Input      - oidsSplTop (the new IdsSplTop property values)
10   Set mIdsSplTop = oidsSplTop
End Property
'*****************************************************
Public Property Get IdsSplTop() As clsIds
Attribute IdsSplTop.VB_Description = "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's up-movement"
'*****************************************************
   ' Purpose    - Returns the Id collection of all virtual splitters which are
   '              effected by the virtual splitter's up-movement
10   Set IdsSplTop = mIdsSplTop
End Property
'*****************************************************
Friend Property Let Left(ByVal lngLeft As Long)
'*****************************************************
   ' Purpose    - Sets the distance between the internal left edge of the virtual
   '              splitter and the left edge of the related Control Manager object
   ' Input      - lngLeft (the new Left property value)
   Dim lngRight            As Long
10   lngRight = Right
20   mlngWidth = lngRight - lngLeft
30   mlngXc = CLng((lngLeft + lngRight) / 2)
End Property
'*****************************************************
Public Property Get Left() As Long
Attribute Left.VB_Description = "Returns the distance between the internal left edge of the virtual splitter and the left edge of the related Control Manager object"
'*****************************************************
   ' Purpose    - Returns the distance between the internal left edge of the
   '              virtual splitter and the left edge of the related Control Manager
   '              object
10   Left = CLng(Xc - (mlngWidth / 2))
End Property
'*****************************************************
Public Property Let LiveUpdate(ByVal blnLiveUpdate As Boolean)
Attribute LiveUpdate.VB_Description = "Returns/sets a value that determines whether the controls should be resized as the splitter is moved"
'*****************************************************
   ' Purpose    - Sets a value that determines whether the controls should be
   '              resized as the splitter is moved
   ' Input      - blnLiveUpdate (the new LiveUpdate property value)
10   mblnLiveUpdate = blnLiveUpdate
End Property
'*****************************************************
Public Property Get LiveUpdate() As Boolean
'*****************************************************
   ' Purpose    - Returns a value that determines whether the controls should be
   '              resized as the splitter is moved
10   LiveUpdate = mblnLiveUpdate
End Property
'*****************************************************
Friend Property Let MaxXc(ByVal lngMaxXc As Long)
'*****************************************************
   ' Purpose    - Sets the maximum x-coordinate of the virtual splitter
   ' Input      - lngMaxXc (the new MaxXc property value)
10   mlngMaxXc = lngMaxXc
End Property
'*****************************************************
Public Property Get MaxXc() As Long
Attribute MaxXc.VB_Description = "Returns the maximum x-coordinate of the virtual splitter"
'*****************************************************
   ' Purpose    - Returns the maximum x-coordinate of the virtual splitter
10   MaxXc = mlngMaxXc
End Property
'*****************************************************
Friend Property Let MaxYc(ByRef lngMaxYc As Long)
'*****************************************************
   ' Purpose    - Sets the maximum y-coordinate of the virtual splitter
   ' Input      - lngMaxYc (the new MaxYc property value)
10   mlngMaxYc = lngMaxYc
End Property
'*****************************************************
Public Property Get MaxYc() As Long
Attribute MaxYc.VB_Description = "Returns the maximum y-coordinate of the virtual splitter"
'*****************************************************
   ' Purpose    - Returns the maximum y-coordinate of the virtual splitter
10   MaxYc = mlngMaxYc
End Property
'*****************************************************
Friend Property Let MinXc(ByRef lngMinXc As Long)
'*****************************************************
   ' Purpose    - Sets the minimum x-coordinate of the virtual splitter
   ' Input      - lngMinXc (the new MinXc property value)
10   mlngMinXc = lngMinXc
End Property
'*****************************************************
Public Property Get MinXc() As Long
Attribute MinXc.VB_Description = "Returns the minimum x-coordinate of the virtual splitter"
'*****************************************************
   ' Purpose    - Returns the minimum x-coordinate of the virtual splitter
10   MinXc = mlngMinXc
End Property
'*****************************************************
Friend Property Let MinYc(ByRef lngMinYc As Long)
'*****************************************************
   ' Purpose    - Sets the minimum y-coordinate of the virtual splitter
   ' Input      - lngMinYc (the new MinYc property value)
10   mlngMinYc = lngMinYc
End Property
'*****************************************************
Public Property Get MinYc() As Long
Attribute MinYc.VB_Description = "Returns the minimum y-coordinate of the virtual splitter"
'*****************************************************
   ' Purpose    - Returns the minimum y-coordinate of the virtual splitter
10   MinYc = mlngMinYc
End Property
'*****************************************************
Friend Property Let Orientation(ByVal udeOrientation As genmSplitterOrientation)
'*****************************************************
   ' Purpose    - Sets the virtual splitter movement direction
   ' Input      - udeOrientation (the new Orientation property value)
10   mudeOrientation = udeOrientation
End Property
'*****************************************************
Public Property Get Orientation() As genmSplitterOrientation
Attribute Orientation.VB_Description = "Returns the virtual splitter movement direction"
'*****************************************************
   ' Purpose    - Returns the virtual splitter movement direction
10   Orientation = mudeOrientation
End Property
'*****************************************************
Friend Property Set Parent(ByRef splsParent As clsSplitters)
'*****************************************************
   ' Purpose    - Sets the collection that contain this class
   ' Input      - splsParent (the new Parent property value)
10   Set mParent = splsParent
End Property
'*****************************************************
Friend Property Get Parent() As clsSplitters
Attribute Parent.VB_Description = "Returns the collection that contain this class"
'*****************************************************
   ' Purpose    - Returns the collection that contain this class
10   Set Parent = mParent
End Property
'*****************************************************
Friend Property Let Right(ByVal lngRight As Long)
'*****************************************************
   ' Purpose    - Sets the distance between the internal right edge of the virtual
   '              splitter and the left edge of the related Control Manager object
   ' Input      - lngLeft (the new Left property value)
   Dim lngLeft             As Long
10   lngLeft = Left
20   mlngWidth = lngRight - lngLeft
30   mlngXc = CLng((lngLeft + lngRight) / 2)
End Property
'*****************************************************
Public Property Get Right() As Long
Attribute Right.VB_Description = "Returns the distance between the internal right edge of the virtual splitter and the left edge of the related Control Manager object"
'*****************************************************
   ' Purpose    - Sets the distance between the internal right edge of the virtual
   '              splitter and the left edge of the related Control Manager object
10   Right = CLng(mlngXc + (mlngWidth / 2))
End Property
'*****************************************************
Friend Property Let Top(ByVal lngTop As Long)
'*****************************************************
   ' Purpose    - Sets the distance between the internal top edge of the virtual
   '              splitter and the top edge of the related Control Manager object
   ' Input      - lngLeft (the new Left property value)
   Dim lngBottom           As Long
10   lngBottom = Bottom
20   mlngHeight = lngBottom - lngTop
30   mlngYc = CLng((lngTop + lngBottom) / 2)
End Property
'*****************************************************
Public Property Get Top() As Long
Attribute Top.VB_Description = "Returns the distance between the internal top edge of the virtual splitter and the top edge of the related Control Manager object"
'*****************************************************
   ' Purpose    - Returns the distance between the internal top edge of the virtual
   '              splitter and the top edge of the related Control Manager object
10   Top = CLng(mlngYc - (mlngHeight / 2))
End Property
'*****************************************************
Friend Property Let Width(ByVal lngWidth As Long)
'*****************************************************
   ' Purpose    - Sets the width of the virtual splitter
   ' Input      - lngWidth (the new Width property value)
10   mlngWidth = lngWidth
End Property
'*****************************************************
Public Property Get Width() As Long
Attribute Width.VB_Description = "Returns the width of the virtual splitter"
'*****************************************************
   ' Purpose    - Returns the width of the virtual splitter
10   Width = mlngWidth
End Property
'*****************************************************
Friend Property Let Xc(ByVal lngXc As Long)
'*****************************************************
   ' Purpose    - Sets the x-coordinate of the virtual splitter center
   ' Input      - lngXc (the new Xc property value)
10   mlngXc = lngXc
End Property
'*****************************************************
Public Property Get Xc() As Long
Attribute Xc.VB_Description = "Returns the x-coordinate of the virtual splitter center"
'*****************************************************
   ' Purpose    - Returns the x-coordinate of the virtual splitter center
10   Xc = mlngXc
End Property
'*****************************************************
Friend Property Let Yc(ByVal lngYc As Long)
'*****************************************************
   ' Purpose    - Sets the y-coordinate of the virtual splitter center
   ' Input      - lngYc (the new Yc property value)
10   mlngYc = lngYc
End Property
'*****************************************************
Public Property Get Yc() As Long
Attribute Yc.VB_Description = "Returns the y-coordinate of the virtual splitter center"
'*****************************************************
   ' Purpose    - Returns the y-coordinate of the virtual splitter center
10   Yc = mlngYc
End Property

' Yorgi's 4Matz [Feb 28,2007 23:58:49] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

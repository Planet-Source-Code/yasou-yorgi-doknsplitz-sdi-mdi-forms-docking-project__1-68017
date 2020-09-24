VERSION 1.0 CLASS
BEGIN
MultiUse = -1                                     'True
Persistable = 0                                   'NotPersistable
DataBindingBehavior = 0                           'vbNone
DataSourceBehavior  = 0                           'vbNone
MTSTransactionMode  = 0                           'NotAnMTSObject
END
Attribute VB_Name = "clsControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A class module to implement custom collection of class clsControl"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************************
'** File Name     - clsControls.cls                                           **
'** Description   - A class module to implement custom collection of class    **
'**                 clsControl                                                **
'** Usage         - Instead of using a real control (like TextBox control)    **
'**                 properties directly, VB Control Manager control saves the **
'**                 necessary properties plus several custom properties in a  **
'**                 virtual control. Furthermore, VB Control Manager control  **
'**                 works with this virtual control (next, will be called     **
'**                 control) to manipulate its position and size and then     **
'**                 write it back to the related real control.                **
'** Dependencies  - clsControl, mdlGeneral                                    **
'** Public Members-                                                           **
'**   * Collections- -                                                        **
'**   * Object     - Item (def. r/o)                                          **
'**   * Property   - Count (r/o)                                              **
'**   * Method     - IsExist                                                  **
'**   * Events     - TitleBarCloseVisibleChanged, TitleBarVisibleChange       **
'** Friend Members-                                                           **
'**   * Collections- -                                                        **
'**   * Object     - Item (def. r/o)                                          **
'**   * Properties - Bottom (r/w), Count (r/o),                               **
'**                  DefaultTitleBar_CloseVisible (r/o),                      **
'**                  DefaultTitleBar_Height (r/o),                            **
'**                  DefaultTitleBar_Visible (r/o),                           **
'**                  Height (r/w), Left (r/w), Right (r/w),                   **
'**                  TitleBar_CloseVisible (r/o), TitleBar_Height (r/w),      **
'**                  TitleBar_Visible (r/w), Top (r/w), Width (r/w)           **
'**   * Methods    - Add, Backup, Compact, IsExist, IsValid, Remove,          **
'**                  RemoveHeaps, RemoveHoles, Restore, Stretch               **
'** Notes         - * Height properties value are generated from Bottom and   **
'**                   Top properties value                                    **
'**                 * Width properties value are generated from Left and      **
'**                   Right properties value                                  **
'** Last modified on September 3, 2003                                        **
'*******************************************************************************
Option Explicit
Private Const mconModuleName           As String = "clsControls"
'--- Public Type Declaration
Public Enum genmRemoveHeapDirection
   rhdAllDirection
   rhdHorizontal
   rhdVertical
End Enum
'--- Custom Collection Class Variable
Private mcolControls                   As Collection
'--- Properties Variables
' Developer's Controls Frame Area Position and Size
' Note- These properties below defined the developer's frame of controls
'       collection. It is named developer's frame because the frame area is
'       determined by the developer. After calling Stretch method, the current
'       frame will have the same size of area with the developer's frame's
Private mlngLeft                       As Long
Private mlngTop                        As Long
Private mlngRight                      As Long
Private mlngBottom                     As Long
' Current Controls Frame Area Position and Size
' Note- These properties below defined the current frame of controls collection.
'       Current frame is the minimum frame area needed to cover all control in
'       controls collection. The frame area is automatically adjusted every time
'       a new control is added in controls collection.
Private mlngCurLeft                    As Long
Private mlngCurTop                     As Long
Private mlngCurRight                   As Long
Private mlngCurBottom                  As Long
' Title Bar
Private mblnTitleBar_CloseVisible      As Boolean
Private mlngTitleBar_Height            As Long
Private mblnTitleBar_Visible           As Boolean
Private mlngTitleBar_TBarType          As TBarTypes
Private mlngTitleBar_Position          As TBarOrientation
'--- Property Default Values
Private Const mconDefaultTitleBar_CloseVisible As Boolean = True
Private Const mconDefaultTitleBar_Height As Long = 15
Private Const mconDefaultTitleBar_Visible As Boolean = True
Private Const mconDefaultTitleBar_TBarType As Long = TBT_DEFAULT
Private Const mconDefaultTitleBar_Position As Long = TBO_HORIZONTAL
'--- Other Variables
Private mblnEventDisable               As Boolean 'indicating whether the class
'event shouldn't be triggered
Private murecControlBackup()           As RECT    'backup of the Controls collection's
'                  size and position
'Description- Occurs when the TitleBar_CloseVisible property of the class is
'             changed
'Arguments  - IdControl (a value that uniquely identifies a control)
Public Event TitleBarCloseVisibleChange(sIdControl As String)
'Description- Occurs when the TitleBar_Visible property of the class is changed
'Arguments  - IdControl (a value that uniquely identifies a control)
Public Event TitleBarVisibleChange(sIdControl As String)
'Description- Occurs when the TitleBar_Type property of the class is changed
'Arguments  - IdControl (a value that uniquely identifies a control)
Public Event TitleBarTypeChange(sIdControl As String)
'*****************************************************
Friend Function Add(cctl As Control, Optional octlNew As clsControl) As Boolean
'*****************************************************
   ' Purpose    - Adds a new item to the collection
   ' Assumption - Controls with index IdCtl exists in cctl collection
   ' Effect     - * If control cctl doesn't have Left, Top, Width or Height
   '                property in run-time mode, then this method doesn't have any
   '                effect, i.e. no item will be added to the collection
   '              * Otherwise, the new item (which is a virtual control of
   '                control cctl has been added to the collection and its
   '                position, size and minimum height/width properties have been
   '                initialized
   ' Inputs     - * cctl (controls collection contained in DoknSplitz
   '                      control)
   '              * IdCtl (Index of cctl that will be added to the collection)
   ' Note       - To get the minimum height and width of control cctl, I set
   '              its height and width to 0 (if the minimum height/width of the
   '              control is larger than 0, control cctl will automatically
   '              adjust its height/value to its minimum value), save control
   '              cctl's height and width in the control's MinWidth and
   '              MinHeight property and then restore control cctl's height
   '              and width. Note that if control cctl is visible when you
   '              add it to the collection, it may produce a flickering effect.
   Dim bctlIsVisible       As Boolean            'save the current state of the control
   Dim lngHeightSave       As Long               'to restore control cctl height
   Dim lngWidthSave        As Long               'to restore control cctl width
   Dim sCtlName            As String
10   On Error GoTo Add_Err
20   mblnEventDisable = True
30   Set octlNew = New clsControl
40   If Not IsArray(cctl) Then
50      sCtlName = cctl.Name
60   Else
70      sCtlName = cctl.Name & CStr(cctl.Index)   'append the index so name is unique
80      End If
90   With octlNew
100      Set .Parent = Me
110      .refCtlObj = cctl                        ' reference the control
120      .Key = sCtlName                          'save the Control Name as the key
130      On Error GoTo Add_Err
140      .TitleBar_CloseVisible = mblnTitleBar_CloseVisible
150      .TitleBar_Height = mlngTitleBar_Height
160      .TitleBar_Visible = mblnTitleBar_Visible
170      .TitleBar_TBarType = mlngTitleBar_TBarType
180      .TitleBar_Position = mlngTitleBar_Position
190      End With
200   mblnEventDisable = False
      ' Sets the item's position and size
      '-- Adds to the collection
210   mcolControls.Add octlNew, octlNew.Key
220   SetFrameDimensions octlNew
230   Add = True
240   Add_Exit:
250   On Error GoTo 0
260   Exit Function
270   Add_Err:
280   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", Add", mconModuleName
290   Set octlNew = Nothing
300   Resume Add_Exit
End Function
'*****************************************************
Friend Sub Backup()
'*****************************************************
   ' Purpose    - Backup all members size and position
   Dim lIdx                As Long
   Dim lTot                As Long
   Dim octl                As clsControl         'for enumerating all virtual controls in Controls collection
10   lTot = mcolControls.Count
20   If lTot > 0 Then
30      ReDim murecControlBackup(1 To lTot)
40      For lIdx = 1 To lTot
50         Set octl = mcolControls(lIdx)
60         If Not octl.Closed Then
70            With murecControlBackup(lIdx)
80               .Top = octl.Top
90               .Right = octl.Right
100               .Bottom = octl.Bottom
110               .Left = octl.Left
120               End With
130            End If
140         Next
150      End If
End Sub
'*****************************************************
Friend Property Get Bottom() As Long
Attribute Bottom.VB_Description = "Returns the distance between the bottom edge of the developer's controls frame area and the top edge of Splitter control"
'*****************************************************
   ' Purpose    - Returns the distance between the bottom edge of the developer's
   '              controls frame area and the top edge of DoknSplitz
   '              control
10   Bottom = mlngBottom
End Property
'*****************************************************
Friend Property Let Bottom(ByVal lngBottom As Long)
'*****************************************************
   ' Purpose    - Sets the distance between the bottom edge of the developer's
   '              controls frame area and the top edge of DoknSplitz
   '              control
   ' Input      - lngBottom (the new Bottom propety value)
10   mlngBottom = lngBottom
End Property
'*****************************************************
Private Sub Class_Initialize()
'*****************************************************
10   Set mcolControls = New Collection
20   InitFrameVars
30   mblnTitleBar_CloseVisible = mconDefaultTitleBar_CloseVisible
40   mlngTitleBar_Height = mconDefaultTitleBar_Height * Screen.TwipsPerPixelY
50   mblnTitleBar_Visible = mconDefaultTitleBar_Visible
60   mblnEventDisable = False
End Sub
'*****************************************************
Private Sub Class_Terminate()
'*****************************************************
10   Set mcolControls = Nothing
End Sub
'*****************************************************
Friend Sub Compact()
Attribute Compact.VB_Description = "Compacts each item in the collection so there won't be any space left"
'*****************************************************
   ' Purpose    - Compacts each item in the collection so there won't be any space left
   ' Note       - Even after calling this method, it may still left special spaces
    '              called "holes" that will be removed with RemoveHoles method. See
    '              VB Control Manager control's documentation or help for the
    '              definition of "hole".
   Dim lngXYMeet           As Long               'the x- or y-coordinate where the edges of two
   Dim octl                As clsControl         'for enumerating all items in the collection
10   On Error GoTo Compact_Err
     '-----------------------------------------
     '-- Clear all friend assignments for each control
     '-----------------------------------------
20   For Each octl In mcolControls
30      octl.IdCtlFriendTop = vbNullString
40      octl.IdCtlFriendRight = vbNullString
50      octl.IdCtlFriendBottom = vbNullString
60      octl.IdCtlFriendLeft = vbNullString
70      Next
80   For Each octl In mcolControls
90      If Not octl.Closed Then
100         With octl
               '-- Compact control octl's top
               ' Sets top-side friend control
110            SetFriend .Key
120            If LenB(.IdCtlFriendTop) = 0 Then
                  ' Control octl doesn't have top-side friend, so set the control's top
                  '   to current frame area's top
130               .Top = mlngCurTop
140            Else
                  ' Determines lngXYMeet value- where control octl and its top-side
                  '   friend will be compacted to
150               SetFriend .IdCtlFriendTop
160               lngXYMeet = (.Top + Item(.IdCtlFriendTop).Bottom) \ 2
170               If Item(.IdCtlFriendTop).IdCtlFriendBottom <> .Key Then
180                  lngXYMeet = IIf(lngXYMeet < Item(Item(.IdCtlFriendTop).IdCtlFriendBottom).Top, lngXYMeet, Item(Item(.IdCtlFriendTop).IdCtlFriendBottom).Top)
190                  End If
                  ' Compacts control octl's top and its top-side friend's bottom
200               .Top = lngXYMeet
210               Item(.IdCtlFriendTop).Bottom = lngXYMeet
220               End If
               '-- Compact control octl's right
               ' Set right-side friend control (must do SetFriend again to compensate
               '   the right-side friend change possibility because of previous
               '   compact)
230            SetFriend .Key
240            If LenB(.IdCtlFriendRight) = 0 Then
                  ' Control octl doesn't have eight-side friend, so set the control's
                  '   right to current frame area's right
250               .Right = mlngCurRight
260            Else
                  ' Determine lngXYMeet value- where control octl and its right-side
                  '   friend will be compacted to
270               SetFriend .IdCtlFriendRight
280               lngXYMeet = (.Right + Item(.IdCtlFriendRight).Left) \ 2
290               If Item(.IdCtlFriendRight).IdCtlFriendLeft <> .Key Then lngXYMeet = IIf(lngXYMeet > Item(Item(.IdCtlFriendRight).IdCtlFriendLeft).Right, lngXYMeet, Item(Item(.IdCtlFriendRight).IdCtlFriendLeft).Right)
                  ' Compacts control octl's right and its right-side friend's left
300               .Right = lngXYMeet
310               Item(.IdCtlFriendRight).Left = lngXYMeet
320               End If
               '-- Compact control octl's bottom
               ' Sets bottom-side friend control (must do SetFriend again to compensate
               '   the bottom-side friend change possibility because of previous
               '   compact)
330            SetFriend .Key
340            If LenB(.IdCtlFriendBottom) = 0 Then
                  ' Control octl doesn't have bottom-side friend, so set the control's
                  '   bottom to current frame area's bottom
350               .Bottom = mlngCurBottom
360            Else
                  ' Determines lngXYMeet value- where control octl and its bottom-side
                  '   friend will be compacted to
370               SetFriend .IdCtlFriendBottom
380               lngXYMeet = (.Bottom + Item(.IdCtlFriendBottom).Top) \ 2
390               If Item(.IdCtlFriendBottom).IdCtlFriendTop <> .Key Then lngXYMeet = IIf(lngXYMeet > Item(Item(.IdCtlFriendBottom).IdCtlFriendTop).Bottom, lngXYMeet, Item(Item(.IdCtlFriendBottom).IdCtlFriendTop).Bottom)
                  ' Compacts control octl's bottom and its bottom-side friend's top
400               .Bottom = lngXYMeet
410               Item(.IdCtlFriendBottom).Top = lngXYMeet
420               End If
               '-- Compact control octl's left
               ' Sets left-side friend control (must do SetFriend again to compensate
               '   the left-side friend change possibility because of previous compact)
430            SetFriend .Key
440            If LenB(.IdCtlFriendLeft) = 0 Then
                  ' Control octl doesn't have left-side friend, so set the control's
                  '   left to current frame area's left
450               .Left = mlngCurLeft
460            Else
                  ' Determines lngXYMeet value- where control octl and its left-side
                  '   friend will be compacted to
470               SetFriend .IdCtlFriendLeft
480               lngXYMeet = (.Left + Item(.IdCtlFriendLeft).Right) \ 2
490               If Item(.IdCtlFriendLeft).IdCtlFriendRight <> .Key Then lngXYMeet = IIf(lngXYMeet < Item(Item(.IdCtlFriendLeft).IdCtlFriendRight).Left, lngXYMeet, Item(Item(.IdCtlFriendLeft).IdCtlFriendRight).Left)
                  ' Compacts control octl's left and its left-side friend's right
500               .Left = lngXYMeet
510               Item(.IdCtlFriendLeft).Right = lngXYMeet
520               End If
530            End With
540         End If
550      Next
560   Compact_Exit:
570   On Error GoTo 0
580   Exit Sub
590   Compact_Err:
#If DebugMode Then
600   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", Compact", mconModuleName
#End If
610   Resume Compact_Exit
End Sub
'*****************************************************
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the collection"
'*****************************************************
   ' Purpose    - Returns the number of items in the collection
10   Count = mcolControls.Count
End Property
'*****************************************************
Friend Property Get DefaultTitleBar_CloseVisible() As Boolean
'*****************************************************
   ' Purpose    - Returns the default TitleBar_CloseVisible property
10   DefaultTitleBar_CloseVisible = mconDefaultTitleBar_CloseVisible
End Property
'*****************************************************
Friend Property Get DefaultTitleBar_Height() As Long
'*****************************************************
   ' Purpose    - Returns the default TitleBar_Height property
10   DefaultTitleBar_Height = mconDefaultTitleBar_Height * Screen.TwipsPerPixelY
End Property
'*****************************************************
Friend Property Get DefaultTitleBar_Position() As TBarOrientation
'*****************************************************
   ' Purpose    - Returns the default TitleBar_Position property
10   DefaultTitleBar_Position = mconDefaultTitleBar_Position
End Property
'*****************************************************
Friend Property Get DefaultTitleBar_TBarType() As TBarTypes
'*****************************************************
   ' Purpose    - Returns the default TitleBar_TBarType property
10   DefaultTitleBar_TBarType = mconDefaultTitleBar_TBarType
End Property
'*****************************************************
Friend Property Get DefaultTitleBar_Visible() As Boolean
'*****************************************************
   ' Purpose    - Returns the default TitleBar_Visible property
10   DefaultTitleBar_Visible = mconDefaultTitleBar_Visible
End Property
'*****************************************************
Public Sub DumpCtls()
'*****************************************************
   Dim lIdx                As Long
   Dim lTot                As Long
   Dim octl                As clsControl         'for enumerating all virtual controls in Controls collection
#If DebugMode Then
10   lTot = mcolControls.Count
20   For lIdx = 1 To lTot
30      Set octl = mcolControls(lIdx)
40      AppTrace mconModuleName, "DumpCtls", "mcolControls ..vctl(" & lIdx & ") name:" & octl.Key
50      Next
#End If
End Sub
'*****************************************************
Friend Property Let Height(ByVal lngHeight As Long)
'*****************************************************
   ' Purpose    - Sets the height of the developer's controls frame area
   ' Input      - lngHeight(the new Height property value)
10   mlngBottom = mlngTop + lngHeight
20   mlngCurBottom = mlngBottom
End Property
'*****************************************************
Friend Property Get Height() As Long
Attribute Height.VB_Description = "Returns the height of the developer's controls frame area"
'*****************************************************
   ' Purpose    - Returns the height of the developer's controls frame area
10   Height = mlngBottom - mlngTop
End Property
'*****************************************************
Private Sub InitFrameVars()
'*****************************************************
   ' Initialize the left and top to the maximum value and the right and bottom to
   '   the minimum value to make sure that these values will be replaced with
   '   minimum left and top value and maximum right and bottom value
10   mlngCurLeft = gconLngInfinite
20   mlngCurTop = gconLngInfinite
30   mlngCurRight = 0
40   mlngCurBottom = 0
End Sub
'*****************************************************
Private Function IsEmptyHole(ByVal octlHole As clsControl) As Boolean
Attribute IsEmptyHole.VB_Description = "Returns the value indicating whether the hole octlHole is empty"
'*****************************************************
   ' Purpose    - Returns the value indicating whether the hole octlHole is empty
   ' Assumption - Hole octlHole's friend controls have been initialized
   ' Note       - See VB Control Manager documentation for the explanation of the
   '              hole which is not empty
   Dim blnIsEmptyHole      As Boolean            'returned value
   Dim octl                As clsControl         'for enumerating all items in the collection
10   If (LenB(octlHole.IdCtlFriendTop) = 0) Or (LenB(octlHole.IdCtlFriendRight) = 0) Or (LenB(octlHole.IdCtlFriendBottom) = 0) Or (LenB(octlHole.IdCtlFriendLeft) = 0) Then
20      blnIsEmptyHole = False
30   Else
40      blnIsEmptyHole = True
50      For Each octl In mcolControls
60         If Not octl.Closed Then
70            With octl
80               If (.Top >= Item(octlHole.IdCtlFriendTop).Bottom) And (.Right <= Item(octlHole.IdCtlFriendRight).Left) And (.Bottom <= Item(octlHole.IdCtlFriendBottom).Top) And (.Left >= Item(octlHole.IdCtlFriendLeft).Right) Then
90                  blnIsEmptyHole = False
100                  Exit For
110                  End If
120               End With
130            End If
140         Next
150      End If
160   IsEmptyHole = blnIsEmptyHole
End Function
'*****************************************************
Public Function IsExist(sId As String) As Boolean
Attribute IsExist.VB_Description = "Returns a value indicating whether certain index that uniquely identifies a virtual control exist in the collection"
'*****************************************************
   ' Purpose    - Returns value indicating whether an id is exist in the collection
   ' Input      - Id
10   On Error Resume Next
20   IsExist = IsObject(mcolControls(sId))
30   On Error GoTo 0
End Function
'*****************************************************
Friend Function IsValid() As Boolean
Attribute IsValid.VB_Description = "Returns value indicating whether each item's width and height in the collection are not smaller than their minimum value"
'*****************************************************
   ' Purpose    - Returns value indicating whether each item's width and height in
   '              the collection are not smaller than their minimum value
   Dim blnIsValid          As Boolean            'returned value
   Dim octl                As clsControl         'for enumerating all items in the collection
10   blnIsValid = True
     ''''      If Not mblnEventDisable Then  'allow changing property values thru the gate
20   For Each octl In mcolControls
30      If Not octl.Closed Then
40         If (octl.Width < octl.MinWidth) Or (octl.Height < octl.MinHeight) Then
50            blnIsValid = False
60            Exit For
70            End If
80         End If
90      Next
     ''''End If
100   IsValid = blnIsValid
End Function
'*****************************************************
Public Property Get Item(ByRef sCtlName As String) As clsControl
Attribute Item.VB_UserMemId = 0
'*****************************************************
   ' Purpose    - Returns an item in the collection which has key IdCtl
   ' Assumption - Key IdCtl exists in the collection
   ' Input      - IdCtl
10   Set Item = mcolControls.Item(sCtlName)
End Property
'*****************************************************
Public Property Get ItemNo(ByRef lCtlNo As Long) As clsControl
'*****************************************************
   ' Purpose    - Returns an item in the collection which has key IdCtl
   ' Assumption - Key IdCtl exists in the collection
   ' Input      - IdCtl
10   Set ItemNo = mcolControls.Item(lCtlNo)
End Property
'*****************************************************
Friend Property Get Left() As Long
Attribute Left.VB_Description = "Returns the distance between the left edge of the developer's controls frame area and the left edge of Splitter control"
'*****************************************************
   ' Purpose    - Returns the distance between the left edge of the developer's
   '              controls frame area and the left edge of DoknSplitz
   '              control
10   Left = mlngLeft
End Property
'*****************************************************
Friend Property Let Left(lngLeft As Long)
'*****************************************************
   ' Purpose    - Sets the distance between the left edge of the developer's
   '              controls frame area and the left edge of DoknSplitz
   '              control
   ' Input      - lngLeft (the new Left property value)
10   mlngLeft = lngLeft
End Property
'*****************************************************
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_MemberFlags = "40"
Attribute NewEnum.VB_UserMemId = -4
'*****************************************************
   ' Purpose    - Enables For Each ... Next enumeration
10   Set NewEnum = mcolControls.[_NewEnum]
End Property
'*****************************************************
Public Sub ReCalcAllDimensions() '4Matz:New
'*****************************************************
   Dim octl                As clsControl
10   InitFrameVars
20   For Each octl In mcolControls
30      SetFrameDimensions octl
40      Next
End Sub
'*****************************************************
Friend Sub Remove(sCtlName As String)
Attribute Remove.VB_Description = "Removes an item which has key IdCtl from the collection"
'*****************************************************
   ' Purpose    - Removes an item which has key IdCtl from the collection
   ' Assumption - Key IdCtl exists in the collection
   ' Input      - IdCtl
10   mcolControls.Remove sCtlName
End Sub
'*****************************************************
Friend Sub RemoveHeap(sCtlName As String, Optional ByVal blnMaintainSize As Boolean = False, Optional ByVal udeRemoveDirection As genmRemoveHeapDirection = rhdAllDirection)
'*****************************************************
   ' Purpose    - Removes item IdCtl's heap
   ' Assumption - Item IdCtl exits in the collection
   ' Input      - * IdCtl
   '              * blnMaintainSize (indicating whether the position and size of item IdCtl can be change)
   '              * udeRemoveDirection (indicating the direction in which way the heaps should be removed)
   Dim blnSuccess          As Boolean
   Dim intNError           As Integer
   Dim lngResizeBottom     As Long               'the area lost for resizing the bottom of item IdCtl
   Dim lngResizeLeft       As Long               'the area lost for resizing the left of item IdCtl
   Dim lngResizeMin        As Long               'the minimum area lost for resizing the items
   Dim lngResizeRight      As Long               'the area lost for resizing the right of item IdCtl
   Dim lngResizeTop        As Long               'the area lost for resizing the top of item IdCtl
   Dim octl1               As clsControl         'represents item IdCtl
   Dim octl2               As clsControl         'for enumerating all items in the collection
   Dim octlHeap            As clsControl         'represents the heap area in item octl1 coused by item octl2
10   On Error GoTo RemoveHeap_Err
20   Set octlHeap = New clsControl
30   Set octl1 = Item(sCtlName)
40   For Each octl2 In mcolControls
50      If Not octl2.Closed Then
60         If octl2.Key <> octl1.Key Then
              '-- Determines the heap-candidate position and size
70            octlHeap.Left = IIf(octl1.Left >= octl2.Left, octl1.Left, octl2.Left)
80            octlHeap.Top = IIf(octl1.Top >= octl2.Top, octl1.Top, octl2.Top)
90            octlHeap.Right = IIf(octl1.Right <= octl2.Right, octl1.Right, octl2.Right)
100            octlHeap.Bottom = IIf(octl1.Bottom <= octl2.Bottom, octl1.Bottom, octl2.Bottom)
               '-- If octlHeap is truely a heap, remove the heap from item IdCtl
110            If ((octlHeap.Left < octlHeap.Right) And (octlHeap.Top < octlHeap.Bottom)) Then
                  ' Determines the area lost to resize the item octl1 to certain direction
120               lngResizeTop = Abs(octl1.Top - octl2.Bottom)
130               lngResizeBottom = Abs(octl1.Bottom - octl2.Top)
140               lngResizeRight = Abs(octl1.Right - octl2.Left)
150               lngResizeLeft = Abs(octl1.Left - octl2.Right)
160               Select Case udeRemoveDirection
                     Case rhdHorizontal
170                     lngResizeTop = gconLngInfinite
180                     lngResizeBottom = gconLngInfinite
190                  Case rhdVertical
200                     lngResizeLeft = gconLngInfinite
210                     lngResizeRight = gconLngInfinite
220                  End Select
                  ' Removes the heap using the minimum area lost
230               intNError = 0
240               Do
250                  blnSuccess = True
260                  Select Case GetMin(lngResizeTop, lngResizeRight, lngResizeBottom, lngResizeLeft)
                        Case lngResizeTop
270                        If blnMaintainSize Then
280                           If octl2.Top >= octlHeap.Top Then
290                              blnSuccess = False
300                              lngResizeTop = gconLngInfinite
310                              intNError = intNError + 1
320                           Else
330                              octl2.Bottom = octlHeap.Top
340                              End If
350                        Else
360                           octl1.Top = octlHeap.Yc
370                           octl2.Bottom = octlHeap.Yc
380                           End If
390                     Case lngResizeRight
400                        If blnMaintainSize Then
410                           If octl2.Right <= octlHeap.Right Then
420                              blnSuccess = False
430                              lngResizeRight = gconLngInfinite
440                              intNError = intNError + 1
450                           Else
460                              octl2.Left = octlHeap.Right
470                              End If
480                        Else
490                           octl1.Right = octlHeap.Xc
500                           octl2.Left = octlHeap.Xc
510                           End If
520                     Case lngResizeBottom
530                        If blnMaintainSize Then
540                           If octl2.Bottom <= octlHeap.Bottom Then
550                              blnSuccess = False
560                              lngResizeBottom = gconLngInfinite
570                              intNError = intNError + 1
580                           Else
590                              octl2.Top = octlHeap.Bottom
600                              End If
610                        Else
620                           octl1.Bottom = octlHeap.Yc
630                           octl2.Top = octlHeap.Yc
640                           End If
650                     Case lngResizeLeft
660                        If blnMaintainSize Then
670                           If octl2.Left >= octlHeap.Left Then
680                              blnSuccess = False
690                              lngResizeLeft = gconLngInfinite
700                              intNError = intNError + 1
710                           Else
720                              octl2.Right = octlHeap.Left
730                              End If
740                        Else
750                           octl1.Left = octlHeap.Xc
760                           octl2.Right = octlHeap.Xc
770                           End If
780                     End Select
790                  Loop Until blnSuccess Or (intNError = 4)
800                  End If
810               End If
820            End If
830         Next
840      RemoveHeap_Exit:
850      On Error Resume Next
860      Set octl1 = Nothing
870      Set octlHeap = Nothing
880      On Error GoTo 0
890      Exit Sub
900      RemoveHeap_Err:
910      ErrHandler Err, Error$, "Line:" & VBA.Erl & ", RemoveHeap", mconModuleName
920      Resume RemoveHeap_Exit
End Sub
'*****************************************************
Friend Sub RemoveHeaps()
Attribute RemoveHeaps.VB_Description = "Removes all items' heap "
'*****************************************************
   ' Purpose    - Removes all items' heap
   Dim octl                As clsControl         'for enumerating all items in the collection
10   For Each octl In mcolControls
20      If Not octl.Closed Then RemoveHeap octl.Key
30      Next
End Sub
'*****************************************************
Friend Sub RemoveHoles() '4Matz:Changed
Attribute RemoveHoles.VB_Description = "Removes all ""holes"" in the collection"
'*****************************************************
    ' Purpose    - Removes all "holes" in the collection
    ' Note       - Even after calling this method, the collection may still left
    '              "holes" for particular complex position of the items in the
    '              collection. See VB Control Manager control's documentation or
    '              help for the definition of "hole" and explanation of complex
    '              position which VB Control Manager couldn't handle.
   Dim blnSuccess          As Boolean            'indicating whether the remove action for certain direction is successed
   Dim octl                As clsControl         'for enumerating all items in the collection
   Dim octlHole            As clsControl         'to represent the hole position and size
   Dim sIdTemp             As String         'to re-arange the "anti-clockwise hole"
10   On Error GoTo RemoveHoles_Err
20   SetFriends
30   For Each octl In mcolControls
40      If Not octl.Closed Then
50         Set octlHole = New clsControl
60         With octlHole
              '-- Determines whether item octl is one of four items that construct the hole
70            .IdCtlFriendRight = octl.IdCtlFriendBottom
80            If LenB(.IdCtlFriendRight) Then
90               .IdCtlFriendBottom = Item(.IdCtlFriendRight).IdCtlFriendLeft
100               If LenB(.IdCtlFriendBottom) Then
110                  .IdCtlFriendLeft = Item(.IdCtlFriendBottom).IdCtlFriendTop
120                  If LenB(.IdCtlFriendLeft) Then .IdCtlFriendTop = Item(.IdCtlFriendLeft).IdCtlFriendRight
130                  End If
140               End If
150            If (.IdCtlFriendTop = octl.Key) Then
                  '-- There is a hole and item Octl is one of four items that construct
                  '   the hole. The next step is to try to remove the hole.
                  '- If the hole is "anti-clockwise" then re-arange the four items ' Id
                  '    that consruct the hole.
160               If Item(.IdCtlFriendTop).Left > Item(.IdCtlFriendRight).Left Then
170                  sIdTemp = .IdCtlFriendTop
180                  .IdCtlFriendTop = .IdCtlFriendLeft
190                  .IdCtlFriendLeft = .IdCtlFriendBottom
200                  .IdCtlFriendBottom = .IdCtlFriendRight
210                  .IdCtlFriendRight = sIdTemp
220                  End If
                  '- If the hole is not "empty", don 't remove the hole
230               blnSuccess = False
240               If IsEmptyHole(octlHole) Then
                     '- The hole is empty, so it must be removed
                     ' Try to stretch-down the top-side-hole-item
250                  If Not blnSuccess Then
260                     If Item(.IdCtlFriendTop).Right > Item(.IdCtlFriendBottom).Right Then
270                        If Item(.IdCtlFriendTop).Right = Item(.IdCtlFriendRight).Right Then
280                           Item(.IdCtlFriendTop).Bottom = Item(.IdCtlFriendBottom).Top
290                           Item(.IdCtlFriendRight).Top = Item(.IdCtlFriendBottom).Top
300                           blnSuccess = True
310                           End If
320                     Else
330                        If Item(.IdCtlFriendTop).Left = Item(.IdCtlFriendLeft).Left Then
340                           Item(.IdCtlFriendTop).Bottom = Item(.IdCtlFriendBottom).Top
350                           Item(.IdCtlFriendLeft).Top = Item(.IdCtlFriendBottom).Top
360                           blnSuccess = True
370                           End If
380                        End If
390                     End If
                     ' If it's not successed, try to stretch-left the
                     '   right-side-hole-item
400                  If Not blnSuccess Then
410                     If Item(.IdCtlFriendRight).Bottom > Item(.IdCtlFriendLeft).Bottom Then
420                        If Item(.IdCtlFriendRight).Bottom = Item(.IdCtlFriendBottom).Bottom Then
430                           Item(.IdCtlFriendRight).Left = Item(.IdCtlFriendLeft).Right
440                           Item(.IdCtlFriendBottom).Right = Item(.IdCtlFriendLeft).Right
450                           blnSuccess = True
460                           End If
470                     Else
480                        If Item(.IdCtlFriendRight).Top = Item(.IdCtlFriendTop).Top Then
490                           Item(.IdCtlFriendRight).Left = Item(.IdCtlFriendLeft).Right
500                           Item(.IdCtlFriendTop).Right = Item(.IdCtlFriendLeft).Right
510                           blnSuccess = True
520                           End If
530                        End If
540                     End If
                     ' If it's not successed, try to stretch-up the bottom-side-hole-item
550                  If Not blnSuccess Then
560                     If Item(.IdCtlFriendBottom).Left < Item(.IdCtlFriendTop).Left Then
570                        If Item(.IdCtlFriendBottom).Left = Item(.IdCtlFriendLeft).Left Then
580                           Item(.IdCtlFriendBottom).Top = Item(.IdCtlFriendTop).Bottom
590                           Item(.IdCtlFriendLeft).Bottom = Item(.IdCtlFriendTop).Bottom
600                           blnSuccess = True
610                           End If
620                     Else
630                        If Item(.IdCtlFriendBottom).Right = Item(.IdCtlFriendRight).Right Then
640                           Item(.IdCtlFriendBottom).Top = Item(.IdCtlFriendTop).Bottom
650                           Item(.IdCtlFriendRight).Bottom = Item(.IdCtlFriendTop).Bottom
660                           blnSuccess = True
670                           End If
680                        End If
690                     End If
                     ' If it's not successed, try to stretch-right the
                     '   left-side-hole-item
700                  If Not blnSuccess Then
710                     If Item(.IdCtlFriendLeft).Top < Item(.IdCtlFriendRight).Top Then
720                        If Item(.IdCtlFriendLeft).Top = Item(.IdCtlFriendTop).Top Then
730                           Item(.IdCtlFriendLeft).Right = Item(.IdCtlFriendRight).Left
740                           Item(.IdCtlFriendTop).Left = Item(.IdCtlFriendRight).Left
750                           blnSuccess = True
760                           End If
770                     Else
780                        If Item(.IdCtlFriendLeft).Bottom = Item(.IdCtlFriendBottom).Bottom Then
790                           Item(.IdCtlFriendLeft).Right = Item(.IdCtlFriendRight).Left
800                           Item(.IdCtlFriendBottom).Left = Item(.IdCtlFriendRight).Left
810                           End If
820                        End If
830                     End If
840                  End If
850               End If
860            End With
870         Set octlHole = Nothing
880         End If
890      Next
900   RemoveHoles_Exit:
910   On Error GoTo 0
920   Exit Sub
930   RemoveHoles_Err:
940   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", RemoveHoles", mconModuleName
950   Resume RemoveHoles_Exit
End Sub
'*****************************************************
Friend Sub Restore()
'*****************************************************
   ' Purpose    - Restore all members size and position to their previous backup values
   ' Assumption - Backup sub routine has been called
   '              in Controls collection
   Dim lIdx                As Long
   Dim lTot                As Long
   Dim octl                As clsControl         'for enumerating all virtual controls in Controls collection
10   lTot = mcolControls.Count
20   If lTot = UBound(murecControlBackup) Then
30      For lIdx = 1 To lTot
40         Set octl = mcolControls(lIdx)
50         With octl
60            If Not .Closed Then
70               .Top = murecControlBackup(lIdx).Top
80               .Right = murecControlBackup(lIdx).Right
90               .Bottom = murecControlBackup(lIdx).Bottom
100               .Left = murecControlBackup(lIdx).Left
110               End If
120            End With
130         Next
140      End If
End Sub
'*****************************************************
Friend Property Let Right(ByVal lngRight As Long)
'*****************************************************
   ' Purpose    - Sets the distance between the right edge of the developer's
   '              controls frame area and the left edge of DoknSplitz
   '              control
   ' Input      - lngRight (the new Right property value)
10   mlngRight = lngRight
End Property
'*****************************************************
Friend Property Get Right() As Long
Attribute Right.VB_Description = "Returns the distance between the right edge of the developer's controls frame area and the left edge of Splitter control"
'*****************************************************
   ' Purpose    - Returns the distance between the right edge of the developer's
   '              controls frame area and the left edge of DoknSplitz
   '              control
10   Right = mlngRight
End Property
'*****************************************************
Private Sub SetFrameDimensions(octlNew As clsControl) '4Matz:New
'*****************************************************
10   With octlNew.refCtlObj
20      octlNew.Left = .Left
30      octlNew.Top = .Top
40      octlNew.Width = .Width
50      octlNew.Height = .Height
60      End With
     '-- Adjusts the current controls frame area
70   If octlNew.Left < mlngCurLeft Then mlngCurLeft = octlNew.Left
80   If octlNew.Top < mlngCurTop Then mlngCurTop = octlNew.Top
90   If octlNew.Right > mlngCurRight Then mlngCurRight = octlNew.Right
100   If octlNew.Bottom > mlngCurBottom Then mlngCurBottom = octlNew.Bottom
End Sub
'*****************************************************
Private Sub SetFriend(sCtlName As String)
Attribute SetFriend.VB_Description = "Sets item IdCtl's friends control"
'*****************************************************
   ' Purpose    - Sets item IdCtl's friends control
   ' Assumption - Item IdCtl exits in the collection
   ' Input      - IdCtl
   Dim octl1               As clsControl         'represents item IdCtl
   Dim octl2               As clsControl         'for enumerating all items in the collection
10   Set octl1 = Item(sCtlName)
20   For Each octl2 In mcolControls
30      If Not octl2.Closed Then
40         If octl2.Key <> octl1.Key Then
50            With octl1
                 ' Sets the top-side friend control
60               If (octl2.Bottom <= .Top) And (((.Left < octl2.Left) And (octl2.Left < .Right)) Or ((.Left < octl2.Right) And (octl2.Right < .Right)) Or ((octl2.Left <= .Left) And (.Right <= octl2.Right))) Then
70                  If LenB(.IdCtlFriendTop) = 0 Then
80                     .IdCtlFriendTop = octl2.Key
90                  ElseIf (octl2.Bottom > Item(.IdCtlFriendTop).Bottom) Or ((octl2.Bottom = Item(.IdCtlFriendTop).Bottom) And (octl2.Right > Item(.IdCtlFriendTop).Right)) Then
100                     .IdCtlFriendTop = octl2.Key
110                     End If
120                  End If
                  ' Sets the right-side friend control
130               If (.Right <= octl2.Left) And (((.Top < octl2.Top) And (octl2.Top < .Bottom)) Or ((.Top < octl2.Bottom) And (octl2.Bottom < .Bottom)) Or ((octl2.Top <= .Top) And (.Bottom <= octl2.Bottom))) Then
140                  If LenB(.IdCtlFriendRight) = 0 Then
150                     .IdCtlFriendRight = octl2.Key
160                  ElseIf (octl2.Left < Item(.IdCtlFriendRight).Left) Or ((octl2.Left = Item(.IdCtlFriendRight).Left) And (octl2.Bottom > Item(.IdCtlFriendRight).Bottom)) Then
170                     .IdCtlFriendRight = octl2.Key
180                     End If
190                  End If
                  ' Sets the bottom-side friend control
200               If (.Bottom <= octl2.Top) And (((.Left < octl2.Left) And (octl2.Left < .Right)) Or ((.Left < octl2.Right) And (octl2.Right < .Right)) Or ((octl2.Left <= .Left) And (.Right <= octl2.Right))) Then
210                  If LenB(.IdCtlFriendBottom) = 0 Then
220                     .IdCtlFriendBottom = octl2.Key
230                  ElseIf (octl2.Top < Item(.IdCtlFriendBottom).Top) Or ((octl2.Top = Item(.IdCtlFriendBottom).Top) And (octl2.Left < Item(.IdCtlFriendBottom).Left)) Then
240                     .IdCtlFriendBottom = octl2.Key
250                     End If
260                  End If
                  ' Sets the left-side friend control
270               If (octl2.Right <= .Left) And (((.Top < octl2.Top) And (octl2.Top < .Bottom)) Or ((.Top < octl2.Bottom) And (octl2.Bottom < .Bottom)) Or ((octl2.Top <= .Top) And (.Bottom <= octl2.Bottom))) Then
280                  If LenB(.IdCtlFriendLeft) = 0 Then
290                     .IdCtlFriendLeft = octl2.Key
300                  ElseIf (octl2.Right > Item(.IdCtlFriendLeft).Right) Or ((octl2.Right = Item(.IdCtlFriendLeft).Right) And (octl2.Top < Item(.IdCtlFriendLeft).Top)) Then
310                     .IdCtlFriendLeft = octl2.Key
320                     End If
330                  End If
340               End With
350            End If
360         End If
370      Next
380   Set octl1 = Nothing
End Sub
'*****************************************************
Private Sub SetFriends()
Attribute SetFriends.VB_Description = "Sets all items' friends control"
'*****************************************************
   ' Purpose    - Sets all items' friends control
   Dim octl                As clsControl         'for enumerating all items in the collection
10   For Each octl In mcolControls
20      If Not octl.Closed Then SetFriend octl.Key
30      Next
End Sub
'*****************************************************
Friend Sub Stretch()
Attribute Stretch.VB_Description = "Stretchs the items to fill-up the developer's frame area"
'*****************************************************
   ' Purpose    - Stretchs the items to fill-up the developer's frame area
   ' Assumption - The current frame area's width and height is greater than 0
   ' Effects    - * The current frame area's size is equal with the developer's
   '                frame area's size
   '              * Others, as specified
   '   area and the left edge of developer's frame area
   '   area and the top edge of developer's frame area
   '   relative to developer's frame area
   '    relative to developer's frame area
   Dim lngdXCorner         As Long               'the distance between the left edge of current frame
   Dim lngdYCorner         As Long               'the distance between the top edge of current frame
   Dim octl                As clsControl         'for enumerating all items in the collection
   Dim sngXScale           As Single             'the width scale of current frame area
   Dim sngYScale           As Single             'the height scale of current frame area
10   On Error GoTo Stretch_Err
20   lngdXCorner = mlngCurLeft - mlngLeft
30   lngdYCorner = mlngCurTop - mlngTop
40   sngXScale = Width / (mlngCurRight - mlngCurLeft)
50   sngYScale = Height / (mlngCurBottom - mlngCurTop)
60   For Each octl In mcolControls
70      If Not octl.Closed Then
80         With octl
90            .Left = CLng((.Left - lngdXCorner) * sngXScale)
100            .Top = CLng((.Top - lngdYCorner) * sngYScale)
110            .Right = CLng((.Right - lngdXCorner) * sngXScale)
120            .Bottom = CLng((.Bottom - lngdYCorner) * sngYScale)
130            End With
140         End If
150      Next
160   mlngCurTop = mlngTop
170   mlngCurRight = mlngRight
180   mlngCurBottom = mlngBottom
190   mlngCurLeft = mlngLeft
200   Stretch_Exit:
210   On Error GoTo 0
220   Exit Sub
230   Stretch_Err:
#If DebugMode Then
240   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", Stretch", mconModuleName
#End If
250   Resume Stretch_Exit
End Sub
'*****************************************************
Friend Property Get TitleBar_CloseVisible(Optional sIdControl As String = vbNullString) As Boolean
'*****************************************************
   ' Purpose    - Returns a value that determines whether a close button in the
   '              title bar is visible
   ' Input      - lngIdControl (a value that uniquely identifies a control)
10   TitleBar_CloseVisible = mblnTitleBar_CloseVisible
End Property
'*****************************************************
Friend Property Let TitleBar_CloseVisible(Optional sIdControl As String = vbNullString, ByVal blnTitleBar_CloseVisible As Boolean)
'*****************************************************
   ' Purpose    - Sets a value that determines whether a close button in the title
   '              bar is visible
   ' Inputs     - * sIdControl (a value that uniquely identifies a control)
   '              * blnblnTitleBar_CloseVisible (the new blnTitleBar_CloseVisible
   '                                             property value)
   Dim octl                As clsControl         'for enumerating all items in the collection
10   If LenB(sIdControl) = 0 Then
20      mblnEventDisable = True
30      For Each octl In mcolControls
40         octl.TitleBar_CloseVisible = blnTitleBar_CloseVisible
50         Next
60      mblnEventDisable = False
70      mblnTitleBar_CloseVisible = blnTitleBar_CloseVisible
80   ElseIf Not mblnEventDisable Then
90      RaiseEvent TitleBarCloseVisibleChange(sIdControl)
100      End If
End Property
'*****************************************************
Friend Property Get TitleBar_Height() As Long
'*****************************************************
   ' Purpose    - Returns the height of the title bar that is shown
10   If Not mblnTitleBar_Visible Then
20      TitleBar_Height = 0
30   Else
40      TitleBar_Height = mlngTitleBar_Height
50      End If
End Property
'*****************************************************
Friend Property Let TitleBar_Height(ByVal lngTitleBar_Height As Long)
'*****************************************************
   ' Purpose    - Sets the height of the title bar
   ' Input      - lngTitleBar_Height (the new TitleBar_Height property value)
   Dim octl                As clsControl         'for enumerating all items in the collection
10   For Each octl In mcolControls
20      octl.TitleBar_Height = lngTitleBar_Height
30      Next
40   mlngTitleBar_Height = lngTitleBar_Height
End Property
'*****************************************************
Friend Property Get TitleBar_Position(Optional sIdControl As String = vbNullString) As TBarOrientation
'*****************************************************
   ' Purpose    - Returns a value that determines whether the title bar is visible
   '              in all the controls in the collection
   ' Input      - lngIdControl (a value that uniquely identifies a control)
10   TitleBar_Position = mlngTitleBar_Position
End Property
'*****************************************************
Friend Property Let TitleBar_Position(Optional sIdControl As String = vbNullString, ByVal lTitleBar_Position As TBarOrientation)
'*****************************************************
   ' Purpose    - Sets a value that determines whether the title bar is visible in
   '              all the controls in the collection
   ' Inputs     - * lngIdControl (a value that uniquely identifies a control)
   '              * lTitleBar_Position (the new lTitleBar_Position property
   '                                        value)
   Dim octl                As clsControl         'for enumerating all items in the collection
10   If LenB(sIdControl) = 0 Then
20      mblnEventDisable = True
30      For Each octl In mcolControls
40         octl.TitleBar_Position = lTitleBar_Position
50         Next
60      mblnEventDisable = False
70      mlngTitleBar_Position = lTitleBar_Position
80   ElseIf Not mblnEventDisable Then RaiseEvent TitleBarTypeChange(sIdControl)
90      End If
End Property
'*****************************************************
Friend Property Let TitleBar_TBarType(Optional sIdControl As String = vbNullString, ByVal blnTitleBar_TBarType As TBarTypes)
'*****************************************************
   ' Purpose    - Sets a value that determines whether the title bar is visible in
   '              all the controls in the collection
   ' Inputs     - * lngIdControl (a value that uniquely identifies a control)
   '              * blnblnTitleBar_TBarType (the new blnTitleBar_TBarType property
   '                                        value)
   Dim octl                As clsControl         'for enumerating all items in the collection
10   If LenB(sIdControl) = 0 Then
20      mblnEventDisable = True
30      For Each octl In mcolControls
40         octl.TitleBar_TBarType = blnTitleBar_TBarType
50         Next
60      mblnEventDisable = False
70      mlngTitleBar_TBarType = blnTitleBar_TBarType
80   ElseIf Not mblnEventDisable Then RaiseEvent TitleBarTypeChange(sIdControl)
90      End If
End Property
'*****************************************************
Friend Property Get TitleBar_TBarType(Optional sIdControl As String = vbNullString) As TBarTypes
'*****************************************************
   ' Purpose    - Returns a value that determines whether the title bar is visible
   '              in all the controls in the collection
   ' Input      - lngIdControl (a value that uniquely identifies a control)
10   TitleBar_TBarType = mlngTitleBar_TBarType
End Property
'*****************************************************
Friend Property Get TitleBar_Visible(Optional sIdControl As String = vbNullString) As Boolean
'*****************************************************
   ' Purpose    - Returns a value that determines whether the title bar is visible
   '              in all the controls in the collection
   ' Input      - lngIdControl (a value that uniquely identifies a control)
10   TitleBar_Visible = mblnTitleBar_Visible
End Property
'*****************************************************
Friend Property Let TitleBar_Visible(Optional sIdControl As String = vbNullString, ByVal blnTitleBar_Visible As Boolean)
'*****************************************************
   ' Purpose    - Sets a value that determines whether the title bar is visible in
   '              all the controls in the collection
   ' Inputs     - * lngIdControl (a value that uniquely identifies a control)
   '              * blnblnTitleBar_Visible (the new blnTitleBar_Visible property
   '                                        value)
   Dim octl                As clsControl         'for enumerating all items in the collection
10   If LenB(sIdControl) = 0 Then
20      mblnEventDisable = True
30      For Each octl In mcolControls
40         octl.TitleBar_Visible = blnTitleBar_Visible
50         Next
60      mblnEventDisable = False
70      mblnTitleBar_Visible = blnTitleBar_Visible
80   ElseIf Not mblnEventDisable Then RaiseEvent TitleBarVisibleChange(sIdControl)
90      End If
End Property
'*****************************************************
Friend Property Get TitleBar_VisibleHeight() As Long
'*****************************************************
   ' Purpose    - Returns the height of the title bar
10   TitleBar_VisibleHeight = mlngTitleBar_Height
End Property
'*****************************************************
Friend Property Get Top() As Long
Attribute Top.VB_Description = "Returns the distance between the top edge of the developer's controls frame area and the top edge of Splitter control"
'*****************************************************
   ' Purpose    - Returns the distance between the top edge of the developer's
   '              controls frame area and the top edge of DoknSplitz
   '              control
10   Top = mlngTop
End Property
'*****************************************************
Friend Property Let Top(ByVal lngTop As Long)
'*****************************************************
   ' Purpose    - Sets the distance between the top edge of the developer's
   '              controls frame area and the top edge of DoknSplitz
   '              control
   ' Input      - lngTop (the new Top property value)
10   mlngTop = lngTop
End Property
'*****************************************************
Friend Property Let Width(ByVal lngWidth As Long)
'*****************************************************
   ' Purpose    - Sets the width of the developer's controls frame area
   ' Input      - lngWidth(the new Width property value)
10   mlngRight = mlngLeft + lngWidth
20   mlngCurRight = mlngRight
End Property
'*****************************************************
Friend Property Get Width() As Long
Attribute Width.VB_Description = "Returns the width of the developer's controls frame area"
'*****************************************************
   ' Purpose    - Returns the width of the developer's controls frame area
10   Width = mlngRight - mlngLeft
End Property

' Yorgi's 4Matz [Feb 28,2007 23:58:49] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

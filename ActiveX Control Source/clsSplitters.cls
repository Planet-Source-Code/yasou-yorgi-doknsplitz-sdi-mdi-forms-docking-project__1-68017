VERSION 1.0 CLASS
BEGIN
MultiUse = -1                                     'True
Persistable = 0                                   'NotPersistable
DataBindingBehavior = 0                           'vbNone
DataSourceBehavior  = 0                           'vbNone
MTSTransactionMode  = 0                           'NotAnMTSObject
END
Attribute VB_Name = "clsSplitters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A class module to implement custom collection of class clsSplitter"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************************
'** File Name     : clsSplitters.cls                                          **
'** Description   : A class module to implement custom collection of class    **
'**                 clsSplitter                                               **
'** Usage         : VB Control Manager control uses PictureBox control to for **
'**                 the splitters. Instead of using PictureBox properties     **
'**                 directly, VB Control Manager control saves the necessary  **
'**                 properties plus several custom properties in a virtual    **
'**                 splitter. Furthermore, VB Control Manager control works   **
'**                 with this virtual splitter (will be called splitter next) **
'**                 to manipulate its position and then write it back to the  **
'**                 PictureBox control.                                       **
'** Dependencies  : clsSplitter, mdlGeneral                                   **
'** Public Members:                                                           **
'**   * Collections: -                                                        **
'**   * Object     : Item (def. r/o)                                          **
'**   * Property   : Count (r/o)                                              **
'**   * Method     : IsExist                                                  **
'**   * Events     : BackColorChange                                          **
'** Friend Members:                                                           **
'**   * Collections: -                                                        **
'**   * Object     : Item (def. r/o)                                          **
'**   * Properties : ActiveColor (r/w), BackColor (r/w), Bottom (r/w),        **
'**                  ClipCursor (r/w), Count (r/o), DefaultActiveColor (r/o), **
'**                  DefaultClipCursor (r/o), DefaultEnable (r/o),            **
'**                  DefaultLiveUpdate (r/o), DefaultSize (r/o),              **
'**                  Enable (r/w), Height (r/w), Left (r/w),                  **
'**                  LiveUpdate (r/w), MinimumSize (r/o), Right (r/w),        **
'**                  Size (r/w), Top (r/w), Width (r/w)                       **
'**   * Methods    : Add, Clear, IsExist, Remove                              **
'** Last modified on September 5, 2003                                        **
'*******************************************************************************
Option Explicit
Private Const mconModuleName           As String = "clsSplitters"
'--- Custom Collection Class Variable
Private mcolSplitters                  As Collection
'--- Property Variables
' Property variables which appears in the property page
Private mblnClipCursor                 As Boolean
Private mblnEnable                     As Boolean
Private mblnLiveUpdate                 As Boolean
Private mlngActiveColor                As OLE_COLOR
Private mlngBackColor                  As OLE_COLOR
Private mlngSize                       As Long
' Splitters frame area
Private mlngLeft                       As Long
Private mlngTop                        As Long
Private mlngRight                      As Long
Private mlngBottom                     As Long
'--- Property Default Values
Private Const mconDefaultActiveColor   As Long = vbBlack
Private Const mconDefaultClipCursor    As Boolean = False
Private Const mconDefaultEnable        As Boolean = True
Private Const mconDefaultLiveUpdate    As Boolean = True
Private Const mconDefaultSize          As Long = 4
'--- Other Variables
Private mblnEventDisable               As Boolean 'indicating whether the class
'event shouldn't be triggered
Private mlngMinimumSize                As Long    'the minimum splitters' size
'---------------------
' Class Custom Events
'---------------------
'Description- Occurs when the BackColor property of the class is changed
'Arguments  - IdSplitter (a value that uniquely identifies a splitter)
Public Event BackColorChange(ByVal IdSplitter As Long)
Attribute BackColorChange.VB_Description = "Sets/returns the background color used to display the splitters"
'Description- Occurs when the Enable property of the class is changed
'Arguments  - IdSplitter (a value that uniquely identifies a splitter)
Public Event EnableChange(ByVal IdSplitter As Long)
'*****************************************************
Friend Property Get ActiveColor() As OLE_COLOR
'*****************************************************
   ' Purpose    - Returns the background color used to display the splitters when
   '              the user moves it in none live update mode
10   ActiveColor = mlngActiveColor
End Property
'*****************************************************
Friend Property Let ActiveColor(ByVal lngActiveColor As OLE_COLOR)
'*****************************************************
   ' Purpose    - Sets the background color used to display the splitter when the
   '              user moves it in none live update mode
   ' Input      - lngActiveColor (the new ActiveColor property value)
   Dim ospl                As clsSplitter        'for enumerating all items in the collection
10   For Each ospl In mcolSplitters
20      ospl.ActiveColor = lngActiveColor
30      Next
40   mlngActiveColor = lngActiveColor
End Property
'*****************************************************
Friend Sub Add(ByRef octl As clsControl, ByRef octls As clsControls, ByRef lCIdx As Long)
Attribute Add.VB_Description = "Adds new items to the collection"
'*****************************************************
   ' Purpose    - Adds new items to the collection
   ' Effects    : * If does not exist, the top-, right-, bottom- and left-side
   '                splitters of control octl have been added and their necessary
   '                properties have been initialized
   '              * Control octl have been shrunk to give space for the splitters
   '                around it
   '              * This method may left several splitters' id marked deleted that
   '                should be removed after all the splitters have been added
   ' Inputs     - * octl (the control that trigger the splitters addition)
   '              * octls (the controls collection of control octl)
   '               Splitter 's movement area
   '-- If the splitters for control octl already exist, assigns the splitters to
   '   controls octl's IdSplTop, IdSpRight, IdSplBottom or IdSplLeft property and
   '   adjust the splitters' size
   Dim lngPos              As Long               'a temporary variable used in determining
   Dim ospl                As clsSplitter        'for enumerating all items in the collection
   Dim osplNew             As clsSplitter        'the new item which will be added to the collection
10   On Error GoTo Add_Err
20   For Each ospl In mcolSplitters
30      If (ospl.Orientation = orHorizontal) And (octl.Top = ospl.Yc) Then
40         If Not IsExistControlBetween(octls, octl, ospl) Then
50            octl.IdSplTop = ospl.Id
60            If octl.Left < ospl.Left Then ospl.Left = octl.Left
70            If octl.Right > ospl.Right Then ospl.Right = octl.Right
80            End If
90         End If
100      If (ospl.Orientation = orVertical) And (octl.Right = ospl.Xc) Then
110         If Not IsExistControlBetween(octls, octl, ospl) Then
120            octl.IdSplRight = ospl.Id
130            If octl.Top < ospl.Top Then ospl.Top = octl.Top
140            If octl.Bottom > ospl.Bottom Then ospl.Bottom = octl.Bottom
150            End If
160         End If
170      If (ospl.Orientation = orHorizontal) And (octl.Bottom = ospl.Yc) Then
180         If Not IsExistControlBetween(octls, octl, ospl) Then
190            octl.IdSplBottom = ospl.Id
200            If octl.Left < ospl.Left Then ospl.Left = octl.Left
210            If octl.Right > ospl.Right Then ospl.Right = octl.Right
220            End If
230         End If
240      If (ospl.Orientation = orVertical) And (octl.Left = ospl.Xc) Then
250         If Not IsExistControlBetween(octls, octl, ospl) Then
260            octl.IdSplLeft = ospl.Id
270            If octl.Top < ospl.Top Then ospl.Top = octl.Top
280            If octl.Bottom > ospl.Bottom Then ospl.Bottom = octl.Bottom
290            End If
300         End If
310      Next
      '-- If the splitters for control octl doesn't exist, create a new splitter,
      '   add it to the collection, and assigns it to controls octl's IdSplTop,
      '   IdSplRight, IdSplBottom or IdSplLeft
320   mblnEventDisable = True
330   If (octl.IdSplTop = gconUninitializedLong) And (octl.Top <> mlngTop) Then
340      Set osplNew = New clsSplitter
350      With osplNew
360         Set .Parent = Me
370         .Id = mcolSplitters.Count
380         .Orientation = orHorizontal
390         .Xc = octl.Xc
400         .Yc = octl.Top
410         .Width = octl.Width
420         .Height = mlngSize
430         .ActiveColor = mlngActiveColor
440         .BackColor = mlngBackColor
450         .ClipCursor = mblnClipCursor
460         .Enable = mblnEnable
470         .LiveUpdate = mblnLiveUpdate
480         mcolSplitters.Add osplNew, CStr(.Id)
490         octl.IdSplTop = .Id
500         End With
510      Set osplNew = Nothing
520      End If
530   If (octl.IdSplRight = gconUninitializedLong) And (octl.Right <> mlngRight) Then
540      Set osplNew = New clsSplitter
550      With osplNew
560         Set .Parent = Me
570         .Id = mcolSplitters.Count
580         .Orientation = orVertical
590         .Xc = octl.Right
600         .Yc = octl.Yc
610         .Width = mlngSize
620         .Height = octl.Height
630         .ActiveColor = mlngActiveColor
640         .BackColor = mlngBackColor
650         .ClipCursor = mblnClipCursor
660         .Enable = mblnEnable
670         .LiveUpdate = mblnLiveUpdate
680         mcolSplitters.Add osplNew, CStr(.Id)
690         octl.IdSplRight = .Id
700         End With
710      Set osplNew = Nothing
720      End If
730   If (octl.IdSplBottom = gconUninitializedLong) And (octl.Bottom <> mlngBottom) Then
740      Set osplNew = New clsSplitter
750      With osplNew
760         Set .Parent = Me
770         .Id = mcolSplitters.Count
780         .Orientation = orHorizontal
790         .Xc = octl.Xc
800         .Yc = octl.Bottom
810         .Width = octl.Width
820         .Height = mlngSize
830         .ActiveColor = mlngActiveColor
840         .BackColor = mlngBackColor
850         .ClipCursor = mblnClipCursor
860         .Enable = mblnEnable
870         .LiveUpdate = mblnLiveUpdate
880         mcolSplitters.Add osplNew, CStr(.Id)
890         octl.IdSplBottom = .Id
900         End With
910      Set osplNew = Nothing
920      End If
930   If (octl.IdSplLeft = gconUninitializedLong) And (octl.Left <> mlngLeft) Then
940      Set osplNew = New clsSplitter
950      With osplNew
960         Set .Parent = Me
970         .Id = mcolSplitters.Count
980         .Orientation = orVertical
990         .Xc = octl.Left
1000         .Yc = octl.Yc
1010         .Width = mlngSize
1020         .Height = octl.Height
1030         .ActiveColor = mlngActiveColor
1040         .BackColor = mlngBackColor
1050         .ClipCursor = mblnClipCursor
1060         .Enable = mblnEnable
1070         .LiveUpdate = mblnLiveUpdate
1080         mcolSplitters.Add osplNew, CStr(.Id)
1090         octl.IdSplLeft = .Id
1100         End With
1110      Set osplNew = Nothing
1120      End If
1130   mblnEventDisable = False
       '-- Shrinks the control octl's size to give space for the splitters
1140   octl.Left = octl.Left + IIf(octl.Left <> mlngLeft, mlngSize \ 2, 0)
1150   octl.Top = octl.Top + IIf(octl.Top <> mlngTop, mlngSize \ 2, 0)
1160   octl.Right = octl.Right - IIf(octl.Right <> mlngRight, mlngSize \ 2, 0)
1170   octl.Bottom = octl.Bottom - IIf(octl.Bottom <> mlngBottom, mlngSize \ 2, 0)
       '-- Assigns all controls and splitters' id that effected by splitter
       '   octl.IdSplTop, octl.IdSplRight, octl.IdSplBottom and octl.IdSplLeft to the
       '   related splitter
1180   If octl.IdSplTop <> gconUninitializedLong Then
1190      With Me(octl.IdSplTop)
1200         .IdsCtlBottom.Add lCIdx, octl.Key
1210         .IdsSplBottom.Add octl.IdSplRight
1220         If .IdsSplTop.IsExist(octl.IdSplRight) Then
1230            .IdsSplTop.MarkDeleted octl.IdSplRight
1240            .IdsSplBottom.MarkDeleted octl.IdSplRight
1250            End If
1260         .IdsSplBottom.Add octl.IdSplLeft
1270         If .IdsSplTop.IsExist(octl.IdSplLeft) Then
1280            .IdsSplTop.MarkDeleted octl.IdSplLeft
1290            .IdsSplBottom.MarkDeleted octl.IdSplLeft
1300            End If
1310         lngPos = octl.Bottom - octl.MinHeight
1320         If lngPos < .MaxYc Then
1330            .MaxYc = lngPos
1340            .IdCtlFriendBottom = lCIdx
1350            End If
1360         End With
1370      End If
1380   If octl.IdSplRight <> gconUninitializedLong Then
1390      With Me(octl.IdSplRight)
1400         .IdsCtlLeft.Add lCIdx, octl.Key
1410         .IdsSplLeft.Add octl.IdSplTop
1420         If .IdsSplRight.IsExist(octl.IdSplTop) Then
1430            .IdsSplRight.MarkDeleted octl.IdSplTop
1440            .IdsSplLeft.MarkDeleted octl.IdSplTop
1450            End If
1460         .IdsSplLeft.Add octl.IdSplBottom
1470         If .IdsSplRight.IsExist(octl.IdSplBottom) Then
1480            .IdsSplRight.MarkDeleted octl.IdSplBottom
1490            .IdsSplLeft.MarkDeleted octl.IdSplBottom
1500            End If
1510         lngPos = octl.Left + octl.MinWidth
1520         If lngPos > .MinXc Then
1530            .MinXc = lngPos
1540            .IdCtlFriendLeft = lCIdx
1550            End If
1560         End With
1570      End If
1580   If octl.IdSplBottom <> gconUninitializedLong Then
1590      With Me(octl.IdSplBottom)
1600         .IdsCtlTop.Add lCIdx, octl.Key
1610         .IdsSplTop.Add octl.IdSplRight
1620         If .IdsSplBottom.IsExist(octl.IdSplRight) Then
1630            .IdsSplBottom.MarkDeleted octl.IdSplRight
1640            .IdsSplTop.MarkDeleted octl.IdSplRight
1650            End If
1660         .IdsSplTop.Add octl.IdSplLeft
1670         If .IdsSplBottom.IsExist(octl.IdSplLeft) Then
1680            .IdsSplBottom.MarkDeleted octl.IdSplLeft
1690            .IdsSplTop.MarkDeleted octl.IdSplLeft
1700            End If
1710         lngPos = octl.Top + octl.MinHeight
1720         If lngPos > .MinYc Then
1730            .MinYc = lngPos
1740            .IdCtlFriendTop = lCIdx
1750            End If
1760         End With
1770      End If
1780   If octl.IdSplLeft <> gconUninitializedLong Then
1790      With Me(octl.IdSplLeft)
1800         .IdsCtlRight.Add lCIdx, octl.Key
1810         .IdsSplRight.Add octl.IdSplTop
1820         If .IdsSplLeft.IsExist(octl.IdSplTop) Then
1830            .IdsSplLeft.MarkDeleted octl.IdSplTop
1840            .IdsSplRight.MarkDeleted octl.IdSplTop
1850            End If
1860         .IdsSplRight.Add octl.IdSplBottom
1870         If .IdsSplLeft.IsExist(octl.IdSplBottom) Then
1880            .IdsSplLeft.MarkDeleted octl.IdSplBottom
1890            .IdsSplRight.MarkDeleted octl.IdSplBottom
1900            End If
1910         lngPos = octl.Right - octl.MinWidth
1920         If lngPos < .MaxXc Then
1930            .MaxXc = lngPos
1940            .IdCtlFriendRight = lCIdx
1950            End If
1960         End With
1970      End If
1980   Add_Exit:
1990   On Error GoTo 0
2000   Exit Sub
2010   Add_Err:
2020   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", Add", mconModuleName
2030   Resume Add_Exit
End Sub
'*****************************************************
Friend Property Get BackColor(Optional ByVal lngIdSplitter As Long = gconUninitializedLong) As OLE_COLOR
'*****************************************************
   ' Purpose    - Returns the background color used to display the splitters
   ' Input      - lngIdSplitter (a value that uniquely identifies a splitter)
10   BackColor = mlngBackColor
End Property
'*****************************************************
Friend Property Let BackColor(Optional ByVal lngIdSplitter As Long = gconUninitializedLong, ByVal lngBackColor As OLE_COLOR)
'*****************************************************
   ' Purpose    - Sets the background color used to display the splitters
   ' Inputs     - * lngIdSplitter (a value that uniquely identifies a splitter)
   '              * lngBackColor (the new BackColor property value)
   Dim ospl                As clsSplitter        'for enumerating all items in the collection
10   If lngIdSplitter = gconUninitializedLong Then
20      mblnEventDisable = True
30      For Each ospl In mcolSplitters
40         ospl.BackColor = lngBackColor
50         Next
60      mblnEventDisable = False
70      mlngBackColor = lngBackColor
80   ElseIf Not mblnEventDisable Then
90      RaiseEvent BackColorChange(lngIdSplitter)
100      End If
End Property
'*****************************************************
Friend Property Let Bottom(ByVal lngBottom As Long)
Attribute Bottom.VB_Description = "Returns or sets the distance between the bottom edge of the splitters frame area and the top edge of VB Splitter control"
'*****************************************************
   ' Purpose    - Sets the distance between the bottom edge of the splitters frame
   '              area and the top edge of VB Control Manager control
   ' Input      - lngBottom (the new Bottom propety value)
10   mlngBottom = lngBottom
End Property
'*****************************************************
Friend Property Get Bottom() As Long
'*****************************************************
   ' Purpose    - Returns the distance between the bottom edge of the splitters
   '              frame area and the top edge of VB Control Manager control
10   Bottom = mlngBottom
End Property
'*****************************************************
Private Sub Class_Initialize()
'*****************************************************
10   Set mcolSplitters = New Collection
20   mlngActiveColor = mconDefaultActiveColor
30   mblnClipCursor = mconDefaultClipCursor
40   mblnEnable = mconDefaultEnable
50   mblnLiveUpdate = mconDefaultLiveUpdate
60   mlngMinimumSize = IIf(Screen.TwipsPerPixelX > Screen.TwipsPerPixelY, Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)
70   mlngSize = DefaultSize
80   mblnEventDisable = False
End Sub
'*****************************************************
Private Sub Class_Terminate()
'*****************************************************
10   Set mcolSplitters = Nothing
End Sub
'*****************************************************
Friend Sub Clear()
Attribute Clear.VB_Description = "Removes all items in the collection"
'*****************************************************
   ' Purpose    - Removes all items in the collection
   Dim ospl                As clsSplitter        'for enumerating all items in the collection
10   For Each ospl In mcolSplitters
20      Remove ospl
30      Next
End Sub
'*****************************************************
Friend Property Get ClipCursor() As Boolean
'*****************************************************
   ' Purpose    - Returns a value that determines whether the mouse pointer is
   '              confined to the items' region
   ' Note       - Each item has its own ClipCursor properties. This property only
   '              return the last ClipCursor property value set by the users. If
   '              the users set the property individually for each item, this
   '              property may not return the correct ClipCursor property value.
10   ClipCursor = mblnClipCursor
End Property
'*****************************************************
Friend Property Let ClipCursor(ByVal blnClipCursor As Boolean)
Attribute ClipCursor.VB_Description = "Returns or sets a value that determines whether the mouse pointer is confined to the items' region\r\n"
'*****************************************************
   ' Purpose    - Sets a value that determines whether the mouse pointer is
   '              confined to the items' region
   ' Effects    : * The ClipCursor property of all items in the collection has been
   '                set to blnClipCursor
   '              * As specified
   ' Input      - blnClipCursor (the new ClipCursor property value)
   Dim ospl                As clsSplitter        'for enumerating all items in the collection
10   For Each ospl In mcolSplitters
20      ospl.ClipCursor = blnClipCursor
30      Next
40   mblnClipCursor = blnClipCursor
End Property
'*****************************************************
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the collection"
'*****************************************************
   ' Purpose    - Returns the number of items in the collection
10   Count = mcolSplitters.Count
End Property
'*****************************************************
Friend Property Get DefaultActiveColor() As Boolean
'*****************************************************
   ' Purpose    - Returns the default ActiveColor property
10   DefaultActiveColor = mconDefaultActiveColor
End Property
'*****************************************************
Friend Property Get DefaultClipCursor() As Boolean
'*****************************************************
   ' Purpose    - Returns the default ClipCursor property
10   DefaultClipCursor = mconDefaultClipCursor
End Property
'*****************************************************
Friend Property Get DefaultEnable() As Boolean
'*****************************************************
   ' Purpose    - Returns the default Enable property
10   DefaultEnable = mconDefaultEnable
End Property
'*****************************************************
Friend Property Get DefaultLiveUpdate() As Boolean
'*****************************************************
   ' Purpose    - Returns the default LiveUpdate property
10   DefaultLiveUpdate = mconDefaultLiveUpdate
End Property
'*****************************************************
Friend Property Get DefaultSize() As Long
'*****************************************************
   ' Purpose    - Returns the default Size property
10   DefaultSize = mconDefaultSize * Screen.TwipsPerPixelX
End Property
'*****************************************************
Friend Property Let Enable(Optional ByVal lngIdSplitter As Long = gconUninitializedLong, ByVal blnEnable As Boolean)
Attribute Enable.VB_Description = "Returns or sets a value that determines whether the items are movable"
'*****************************************************
   ' Purpose    - Sets a value that determines whether the splitter is movable
   ' Effects    : * The Enable property of all items in the collection has been
   '                set to blnEnable
   '              * As specified
   ' Inputs     - * lngIdSplitter (a value that uniquely identifies a splitter)
   '              * blnEnable (the new Enable property value)
   Dim ospl                As clsSplitter        'for enumerating all items in the collection
10   If lngIdSplitter = gconUninitializedLong Then
20      mblnEventDisable = True
30      For Each ospl In mcolSplitters
40         ospl.Enable = blnEnable
50         Next
60      mblnEventDisable = False
70      mblnEnable = blnEnable
80   ElseIf Not mblnEventDisable Then
90      RaiseEvent EnableChange(lngIdSplitter)
100      End If
End Property
'*****************************************************
Friend Property Get Enable(Optional ByVal lngIdSplitter As Long = gconUninitializedLong) As Boolean
'*****************************************************
   ' Purpose    - Returns a value that determines whether the items are movable
   ' Note       - Each item has its own Enable properties. This property only
   '              return the last Enable property value set by the users. If the
   '              users set the property individually for each item, this property
   '              may not return the correct Enable property value.
10   Enable = mblnEnable
End Property
'*****************************************************
Friend Property Get Height() As Long
'*****************************************************
   ' Purpose    - Returns the height of the splitters frame area
10   Height = mlngBottom - mlngTop
End Property
'*****************************************************
Friend Property Let Height(ByVal lngHeight As Long)
Attribute Height.VB_Description = "Returns or sets the height of the splitters frame area"
'*****************************************************
   ' Purpose    - Sets the height of the splitters frame area
   ' Input      - lngHeight(the new Height property value)
10   mlngBottom = mlngTop + lngHeight
End Property
'*****************************************************
Public Function IsExist(Id As Long) As Boolean
Attribute IsExist.VB_Description = "Returns a value indicating whether certain index that uniquely identifies a virtual splitter exist in the collection"
'*****************************************************
   ' Purpose    - Returns value indicating whether an id is exist in the collection
   ' Input      - Id
   Dim blnIsExist          As Boolean
   Dim ospl                As clsSplitter
10   blnIsExist = False
20   For Each ospl In mcolSplitters
30      If ospl.Id = Id Then
40         blnIsExist = True
50         Exit For
60         End If
70      Next
80   IsExist = blnIsExist
End Function
'*****************************************************
Private Function IsExistControlBetween(octlsSource As clsControls, octlSource As clsControl, osplSource As clsSplitter)
'*****************************************************
   ' Purpose    - Returns a value indicating whether there is a control between
   '              control octlSource and splitter osplSource in controls collection
   '              octlsSource
   ' Inputs     - * octlsSource
   '              * octlSource
   '              * osplSource
   Dim blnIsExistControlBetween As Boolean       'returned value
   Dim octl                As clsControl         'for enumerating all items in the collection
10   blnIsExistControlBetween = False
20   For Each octl In octlsSource
30      Select Case osplSource.Orientation
           Case orHorizontal
40            If (octl.Top <= osplSource.Top) Then
50               If (octl.Bottom >= osplSource.Bottom) Then
60                  If (((octl.Left >= osplSource.Right) And (octl.Right <= octlSource.Left)) Or ((octl.Left >= octlSource.Right) And (octl.Right <= osplSource.Left))) Then
70                     blnIsExistControlBetween = True
80                     Exit For
90                     End If
100                  End If
110               End If
120         Case orVertical
130            If (octl.Left <= osplSource.Left) Then
140               If (octl.Right >= osplSource.Right) Then
150                  If (((octl.Top >= osplSource.Bottom) And (octl.Bottom <= octlSource.Top)) Or ((octl.Top >= octlSource.Bottom) And (octl.Bottom <= osplSource.Top))) Then
160                     blnIsExistControlBetween = True
170                     Exit For
180                     End If
190                  End If
200               End If
210         End Select
220      Next
230   IsExistControlBetween = blnIsExistControlBetween
End Function
'*****************************************************
Public Property Get Item(IdSpl As Long) As clsSplitter
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns an item in the collection which has key IdSpl"
'*****************************************************
   ' Purpose    - Returns an item in the collection which has key IdSpl
   ' Assumption - Key IdSpl exists in the collection
   ' Input      - IdSpl
10   Set Item = mcolSplitters.Item(CStr(IdSpl))
End Property
'*****************************************************
Friend Property Get Left() As Long
'*****************************************************
   ' Purpose    - Returns the distance between the left edge of the splitters frame
   '              area and the left edge of VB Control Manager control
10   Left = mlngLeft
End Property
'*****************************************************
Friend Property Let Left(ByVal lngLeft As Long)
Attribute Left.VB_Description = "Returns or sets the distance between the left edge of the splitters frame area and the left edge of VB Splitter control"
'*****************************************************
   ' Purpose    - Sets the distance between the left edge of the splitters frame
   '              area and the left edge of VB Control Manager control
   ' Input      - lngLeft (the new Left propety value)
10   mlngLeft = lngLeft
End Property
'*****************************************************
Friend Property Let LiveUpdate(ByVal blnLiveUpdate As Boolean)
Attribute LiveUpdate.VB_Description = "Returns or sets a value that determines whether the controls should be resized as the  item is moved"
'*****************************************************
   ' Purpose    - Sets a value that determines whether the controls should be
   '              resized as the item is moved
   ' Effects    - * The LiveUpdate property of all items in the collection has been
   '                set to blnLiveUpdate
   '              * As specified
   ' Input      - blnLiveUpdate (the new LiveUpdate property value)
   Dim ospl                As clsSplitter        'for enumerating all items in the collection
10   For Each ospl In mcolSplitters
20      ospl.LiveUpdate = blnLiveUpdate
30      Next
40   mblnLiveUpdate = blnLiveUpdate
End Property
'*****************************************************
Friend Property Get LiveUpdate() As Boolean
'*****************************************************
   ' Purpose    - Returns a value that determines whether the controls should be
   '              resized as the item is moved
   ' Note       - Each item has its own LiveUpdate properties. This property only
   '              return the last LiveUpdate property value set by the users. If
   '              the users set the property individually for each item, this
   '              property may not return the correct LiveUpdate property value.
10   LiveUpdate = mblnLiveUpdate
End Property
'*****************************************************
Friend Property Get MinimumSize() As Long
Attribute MinimumSize.VB_Description = "Returns the minimum items' size"
'*****************************************************
   ' Purpose    - Returns the minimum items' size
10   MinimumSize = mlngMinimumSize
End Property
'*****************************************************
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_MemberFlags = "40"
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_Description = "Enables For Each ... Next enumeration"
'*****************************************************
   ' Purpose    - Enables For Each ... Next enumeration
10   Set NewEnum = mcolSplitters.[_NewEnum]
End Property
'*****************************************************
Friend Sub Remove(ByVal IdSpl As Long)
Attribute Remove.VB_Description = "Removes an item which has key IdSpl from the collection"
'*****************************************************
   ' Purpose    - Removes an item which has key IdSpl from the collection
   ' Assumption - Key IdSpl exists in the collection
   ' Input      - IdSpl
10   mcolSplitters.Remove CStr(IdSpl)
End Sub
'*****************************************************
Friend Property Let Right(ByVal lngRight As Long)
Attribute Right.VB_Description = "Returns or sets the distance between the right edge of the splitters frame area and the left edge of VB Splitter control"
'*****************************************************
   ' Purpose    - Sets the distance between the right edge of the splitters frame
   '              area and the left edge of VB Control Manager control
   ' Input      - lngRight (the new Right propety value)
10   mlngRight = lngRight
End Property
'*****************************************************
Friend Property Get Right() As Long
'*****************************************************
   ' Purpose    - Returns the distance between the right edge of the splitters
   '              frame area and the left edge of VB Control Manager control
10   Right = mlngRight
End Property
'*****************************************************
Friend Property Let Size(ByVal lngSize As Long)
Attribute Size.VB_Description = "Returns or sets the items' size"
'*****************************************************
   ' Purpose    - Sets the items' size
   ' Effects    : * The Height or Width property (depend on the Orientation
   '                property) of all items in the collection have been set to
   '                lngSize
   '              * As specified
   ' Input      - lngSize (the new Size property value)
   Dim ospl                As clsSplitter        'for enumerating all items in the collection
10   For Each ospl In mcolSplitters
20      Select Case ospl.Orientation
           Case orHorizontal
30            ospl.Height = lngSize
40         Case orVertical
50            ospl.Width = lngSize
60         End Select
70      Next
80   mlngSize = lngSize
End Property
'*****************************************************
Friend Property Get Size() As Long
'*****************************************************
   ' Purpose    - Returns the items' size
   ' Note       - Each item has its own Size properties. This property only return
   '              the last Size property value set by the users. If the users set
   '              the property individually for each item, this property may not
   '              return the correct Size property value.
10   Size = mlngSize
End Property
'*****************************************************
Friend Property Get Top() As Long
'*****************************************************
   ' Purpose    - Returns the distance between the top edge of the splitters frame
   '              area and the top edge of VB Control Manager control
10   Top = mlngTop
End Property
'*****************************************************
Friend Property Let Top(ByVal lngTop As Long)
Attribute Top.VB_Description = "Returns or sets the distance between the top edge of the splitters frame area and the top edge of VB Splitter control"
'*****************************************************
   ' Purpose    - Sets the distance between the top edge of the splitters frame
   '              area and the top edge of VB Control Manager control
   ' Input      - lngTop (the new Top propety value)
10   mlngTop = lngTop
End Property
'*****************************************************
Friend Property Let Width(ByVal lngWidth As Long)
Attribute Width.VB_Description = "Returns or sets the width of the splitters frame area"
'*****************************************************
   ' Purpose    - Sets the width of the splitters frame area
   ' Input      - lngWidth(the new Width property value)
10   mlngRight = mlngLeft + lngWidth
End Property
'*****************************************************
Friend Property Get Width() As Long
'*****************************************************
   ' Purpose    - Returns the width of the splitters frame area
10   Width = mlngRight - mlngLeft
End Property

' Yorgi's 4Matz [Feb 28,2007 23:58:49] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

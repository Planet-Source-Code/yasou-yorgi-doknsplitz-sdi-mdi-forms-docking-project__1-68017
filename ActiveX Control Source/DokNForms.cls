VERSION 1.0 CLASS
BEGIN
MultiUse = -1                                     'True
Persistable = 0                                   'NotPersistable
DataBindingBehavior = 0                           'vbNone
DataSourceBehavior  = 0                           'vbNone
MTSTransactionMode  = 0                           'NotAnMTSObject
END
Attribute VB_Name = "DokNForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"DockForm"
Attribute VB_Ext_KEY = "Member0" ,"DockForm"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Class:             DokNForms
' Description:       Maintains the Collection of "DokNForm" objects.
'                    The form's hWnd is used as a unique handle for lookups.
Option Explicit
Private Const mconModuleName           As String = "DokNForms"
Private mCol                           As Collection ' local variable to hold collection
'*****************************************************
Friend Function Add(ByRef ofrm As Form, ByRef picHost As PictureBox, ByRef oParenthWnd As Long, Style As eDockStyles, sKey As String, Optional iPos As TBarOrientation = TBO_HORIZONTAL, Optional bHasCloseButton As Boolean = False) As DokNForm
'*****************************************************
   Dim lLastErr            As Long
   Dim oNewDoknForm        As DokNForm
10   On Error Resume Next
20   Set oNewDoknForm = mCol.Item(sKey)
30   On Error GoTo Add_Err
40   If oNewDoknForm Is Nothing Then
        ' create a new object
50      Set oNewDoknForm = New DokNForm
60      mCol.Add oNewDoknForm, sKey
70      End If
80   With oNewDoknForm
        'save the Host Container
90      Set .HostContainer = picHost
        'Save original style attributes to restore during UnDock
100      .FloatingStyle = GetWindowLong(ofrm.hWnd, GWL_STYLE)
110      .FloatingExStyle = GetWindowLong(ofrm.hWnd, GWL_EXSTYLE)
120      .HasCloseButton = bHasCloseButton
130      .DockedForm = ofrm
         ' set the properties passed into the method
140      .Style = Style
150      .State = DS_NotReady
160      .VCtlKey = vbNullString
170      .FloatingWidth = ofrm.Width
180      .FloatingHeight = ofrm.Height
190      .FloatingLeft = (Screen.Width \ 2) - (ofrm.Width \ 2) 'init to the center
200      .FloatingTop = (Screen.Height \ 2) - (ofrm.Height \ 2) 'init to the center
210      If oParenthWnd Then
220         .FormParentHwnd = oParenthWnd
230      Else
240         .FormParentHwnd = GetParent(ofrm.hWnd)
250         End If
260      .Key = sKey
270      .TBarPos = iPos
280      End With
      ' return the object created
290   Set Add = oNewDoknForm
300   Add_Exit:
310   On Error Resume Next
320   Set oNewDoknForm = Nothing
330   On Error GoTo 0
340   Exit Function
350   Add_Err:
360   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", Add", mconModuleName
370   Resume Add_Exit
End Function
'*****************************************************
Private Sub Class_Initialize()
'*****************************************************
   ' creates the collection when this class is created
10   Set mCol = New Collection
End Sub
'*****************************************************
Private Sub Class_Terminate()
'*****************************************************
   ' destroys collection when this class is terminated
10   Set mCol = Nothing
End Sub
'*****************************************************
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the collection"
'*****************************************************
10   Count = mCol.Count
End Property
'*****************************************************
Friend Function HandleExists(hWndA As Long) As Boolean
'*****************************************************
   Dim Index               As Long
10   HandleExists = IndexFromHandle(hWndA, Index)
End Function
'*****************************************************
Private Function IndexFromHandle(ByRef hWndA As Long, ByRef lIdx As Long) As Boolean
'*****************************************************
   'Purpose: Using the form's hwnd as a unique identifier allows a secondary key lookup
   Dim lTot                As Long
10   lTot = mCol.Count
20   For lIdx = 1 To lTot
30      If mCol(lIdx).hWnd = hWndA Then
40         IndexFromHandle = True
50         Exit Function
60         End If
70      Next lIdx
End Function
'*****************************************************
Public Property Get Item(Index As Variant) As DokNForm
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns a docked form from the list"
'*****************************************************
10   On Error Resume Next
20   Set Item = mCol.Item(Index)
End Property
'*****************************************************
Friend Function ItemByHandle(hWndA As Long) As DokNForm
'*****************************************************
   Dim Index               As Long
10   If IndexFromHandle(hWndA, Index) Then
20      Set ItemByHandle = mCol(Index)
30   Else
40      Set ItemByHandle = Nothing
50      End If
End Function
'*****************************************************
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_MemberFlags = "40"
Attribute NewEnum.VB_UserMemId = -4
'*****************************************************
10   Set NewEnum = mCol.[_NewEnum]
End Property
'*****************************************************
Public Sub Remove(Index As Variant)
Attribute Remove.VB_Description = "Removes a item (DockForm) from the collection list"
'*****************************************************
10   mCol.Remove Index
End Sub
'*****************************************************
Friend Sub RemoveByHandle(hWndA As Long)
'*****************************************************
   Dim Index               As Long
10   If IndexFromHandle(hWndA, Index) Then
20      mCol.Remove Index
30      End If
End Sub

' Yorgi's 4Matz [Feb 28,2007 23:58:49] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

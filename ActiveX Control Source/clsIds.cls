VERSION 1.0 CLASS
BEGIN
MultiUse = -1                                     'True
Persistable = 0                                   'NotPersistable
DataBindingBehavior = 0                           'vbNone
DataSourceBehavior  = 0                           'vbNone
MTSTransactionMode  = 0                           'NotAnMTSObject
END
Attribute VB_Name = "clsIds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A class module to implement collection of class clsId"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************************
'** File Name     : clsIds.cls                                                **
'** Description   : A class module to implement custom collection of class    **
'**                 clsId                                                     **
'** Usage         : Represents list of Ids of class clsControl's and          **
'**                 clsSplitter's instances                                   **
'** Dependencies  : clsId, mdlGeneral                                         **
'** Public Members:                                                           **
'**   * Collections: -                                                        **
'**   * Object     : Item (def. r/o)                                          **
'**   * Property   : Count (r/o)                                              **
'**   * Methods    : -                                                        **
'**   * Events     : -                                                        **
'** Friend Members:                                                           **
'**   * Collections: -                                                        **
'**   * Object     : Item (def. r/o)                                          **
'**   * Property   : Count (r/o)                                              **
'**   * Methods    : Add, IsExist, IsMarkDeleted, MarkDeleted, Remove,        **
'**                  RemoveDeleted                                            **
'** Last modified on October 10, 2003                                         **
'*******************************************************************************
Option Explicit
'--- Constant Declaration
Private Const conIdDeleted             As Long = -2 'represent the Id which is marked deleted
'--- Custom Collection Class Variable
Private mcolIds                        As Collection
'*****************************************************
Friend Sub Add(Id As Long, Optional ByRef Key As String = vbNullString)
'*****************************************************
   ' Purpose    - Adds a new item to the collection
   ' Effect     - * If Id = gconUninitializedLong or Key is already marked deleted
   '                then this method doesn't have any effect, i.e. no item will be
   '                added to the collection
   '              * if Key already exist, the item's value will be replaced with Id
   '              * Otherwise, as specified
   ' Inputs     - * Id
   '              * Key (if not passed as parameter, Id will be used as Key)
   Dim oidNew              As clsId              'the new item which will be added to the collection
10   On Error Resume Next
20   If Id <> gconUninitializedLong Then
30      Set oidNew = New clsId
40      oidNew = Id
50      If LenB(Key) = 0 Then
60         mcolIds.Add oidNew, CStr(Id)
70      ElseIf Not IsMarkDeleted(Key) Then
80         mcolIds.Remove Key
90         mcolIds.Add oidNew, Key
100         End If
110      Set oidNew = Nothing
120      End If
End Sub
'*****************************************************
Private Sub Class_Initialize()
'*****************************************************
10   Set mcolIds = New Collection
End Sub
'*****************************************************
Private Sub Class_Terminate()
'*****************************************************
10   Set mcolIds = Nothing
End Sub
'*****************************************************
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the collection"
'*****************************************************
   ' Purpose    - Returns the number of items in the collection
10   Count = mcolIds.Count
End Property
'*****************************************************
Friend Function IsExist(Key As Long) As Boolean
Attribute IsExist.VB_Description = "Returns a value that determines whether the item with key Key exists in the collection"
'*****************************************************
   '---------------
   ' Class Methods
   '---------------
   ' Purpose    - Returns a value that determines whether the item with key Key
   '              exists in the collection
   ' Input      - Key
   ' Note       - Item that is marked deleted is considered not exist
10   IsExist = (Me(Key).Id <> gconUninitializedLong) And Not IsMarkDeleted(Key)
End Function
'*****************************************************
Friend Function IsMarkDeleted(ByVal Key As String) As Boolean
Attribute IsMarkDeleted.VB_Description = "Returns a value that determines whether the item with key Key is marked deleted"
'*****************************************************
   ' Purpose    - Returns a value that determines whether the item with key Key
   '              is marked deleted
   ' Input      - Key
10   On Error Resume Next
20   IsMarkDeleted = (mcolIds(Key).Id = conIdDeleted)
30   On Error GoTo 0
End Function
'*****************************************************
Public Property Get Item(Key As Variant) As clsId
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_Description = "Returns an item in the collection which has key Key"
'*****************************************************
   ' Purpose    - Returns an item in the collection which has key Key
   ' Input      - Key
   ' Return     : * If key Key doesn't exist in the collection, returns item with
   '                Id = gconUninitializedLong
   '              * Otherwise, returns as specified
   Dim oidItem             As clsId
10   On Error GoTo ErrorHandler
20   Set oidItem = New clsId
30   Set oidItem = mcolIds.Item(CStr(Key))
40   Set Item = oidItem
50   Exit Property
60   ErrorHandler:
70   oidItem = gconUninitializedLong
80   Resume Next
End Property
'*****************************************************
Friend Sub MarkDeleted(Key As Long)
Attribute MarkDeleted.VB_Description = "Marks deleted the item with key Key"
'*****************************************************
   ' Purpose    - Marks deleted the item with key Key
   ' Input      - Key
10   Add conIdDeleted, CStr(Key)
End Sub
'*****************************************************
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_MemberFlags = "40"
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_Description = "Enables For Each ... Next enumeration"
'*****************************************************
   ' Purpose    - Enables For Each ... Next enumeration
10   Set NewEnum = mcolIds.[_NewEnum]
End Property
'*****************************************************
Friend Sub Remove(Key As Long)
Attribute Remove.VB_Description = "Removes an item which has key Key from the collection"
'*****************************************************
   ' Purpose    - Removes an item which has key Key from the collection
   ' Assumption - Key Key exists in the collection
   ' Input      - Key
10   mcolIds.Remove CStr(Key)
End Sub
'*****************************************************
Friend Sub RemoveDeleted(lngLastPos As Long)
Attribute RemoveDeleted.VB_Description = "Removes all items in the collection which is marked deleted"
'*****************************************************
   ' Purpose    - Removes all items in the collection which is marked deleted
   ' Input      - lngLastPos (the last position in the collection)
   Dim i                   As Long
10   For i = 0 To lngLastPos
20      If IsMarkDeleted(i) Then mcolIds.Remove CStr(i)
30      Next
End Sub

' Yorgi's 4Matz [Feb 28,2007 23:58:49] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

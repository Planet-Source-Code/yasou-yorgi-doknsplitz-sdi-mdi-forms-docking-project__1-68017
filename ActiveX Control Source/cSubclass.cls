VERSION 1.0 CLASS
BEGIN
MultiUse = -1                                     'True
Persistable = 0                                   'NotPersistable
DataBindingBehavior = 0                           'vbNone
DataSourceBehavior  = 0                           'vbNone
MTSTransactionMode  = 0                           'NotAnMTSObject
END
Attribute VB_Name = "cSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* cSubclass - subclassing class. Use with either the interface class (iSubclass.cls) or the
'*             supplied type library (WinSubHook3.tlb).
'*
'* Note: when used with a form, it's a bad idea to Terminate subclassing from a form's terminate
'*       event. The window will have been destroyed and so prevent the thunk from releasing its
'*       allocated memory. Instead, use the form's unload event to Terminate subclassing when
'*       closing a form.
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
'* v1.1 VirtualAlloc memory to prevent Data Execution Prevention faults on Win64......... 20060324
'* v1.2 Thunk redesigned to handle unsubclassing and memory release...................... 20060325
'*************************************************************************************************
Option Explicit
'-cSublass private declarations-------------------------------------------------------------------
Private Const MSG_ENTRIES              As Long = 32 'Number of msg table entries
Private Const CODE_LEN                 As Long = 240 'Thunk length in bytes
Private Const WNDPROC_OFF              As Long = &H30 'WndProc execution offset
Private Const MEmlEN                   As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1)) 'Bytes to allocate per thunk, data + code + msg tables
Private Const PAGE_RWX                 As Long = &H40 'Allocate executable memory
Private Const MEM_COMMIT               As Long = &H1000 'Commit allocated memory
Private Const GWL_WNDPROC              As Long = -4 'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN             As Long = 1 'Shutdown flag data index
Private Const IDX_HWND                 As Long = 2 'hWnd data index
Private Const IDX_EBMODE               As Long = 3 'EbMode data index
Private Const IDX_CWP                  As Long = 4 'CallWindowProc data index
Private Const IDX_SWL                  As Long = 5 'SetWindowsLong data index
Private Const IDX_FREE                 As Long = 6 'VirtualFree data index
Private Const IDX_OWNER                As Long = 7 'Owner data index
Private Const IDX_WNDPROC              As Long = 8 'Original WndProc data index
Private Const IDX_CALLBACK             As Long = 9 'zWndProc data index
Private Const IDX_BTABLE               As Long = 10 'Before table data index
Private Const IDX_ATABLE               As Long = 11 'After table data index
Private Const IDX_EBX                  As Long = 14 'Data code index
Private z_Code(29)                     As Currency 'Thunk machine-code initialised here
Private z_Data(32)                     As Long    'Array whose data pointer is re-mapped to arbitary memory addresses
Private z_DataDataPtr                  As Long    'Address of z_Data()'s SafeArray data pointer
Private z_DataOrigData                 As Long    'Address of z_Data()'s original data
Private z_hWnds                        As Collection 'hWnd/thunk-address collection
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'-cSubclass code----------------------------------------------------------------------------------
'*****************************************************
Public Sub AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As eMsg, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
'*****************************************************
   'Add the message value to the window handle's specified callback table
    'xTrace "AddMsg - hWnd:" & lng_hWnd & ", uMsg:" & Hex$(uMsg)
10   If When And MSG_BEFORE Then                  'If the message is to be added to the before original WndProc table...
20      zAddMsg lng_hWnd, uMsg, IDX_BTABLE        'Add the message to the before table
30      End If
40   If When And MSG_AFTER Then                   'If message is to be added to the after original WndProc table...
50      zAddMsg lng_hWnd, uMsg, IDX_ATABLE        'Add the message to the after table
60      End If
70   zMap z_DataOrigData                          'Restore z_Data()'s original data pointer
End Sub
'*****************************************************
Public Function CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'*****************************************************
   'Call the original WndProc
10   zMap_hWnd lng_hWnd                           'Map z_Data() to the thunk of the specified window handle
20   CallOrigWndProc = CallWindowProcA(z_Data(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
30   zMap z_DataOrigData                          'Restore z_Data()'s original data pointer
End Function
'*****************************************************
Private Sub Class_Initialize()                                              'Class initialization
'*****************************************************
10   RtlMoveMemory VarPtr(z_DataDataPtr), VarPtrArray(z_Data), 4 'Get the address of z_Data()'s SafeArray header
20   z_DataDataPtr = z_DataDataPtr + 12           'Bump the address to point to the pvData data pointer
30   RtlMoveMemory VarPtr(z_DataOrigData), z_DataDataPtr, 4 'Get the value of z_Data()'s SafeArray pvData data pointer
     'Initialise the machine-code thunk
40   z_Code(6) = -490736517001394.5807@: z_Code(7) = 484417356483292.94@: z_Code(8) = -171798741966746.6996@: z_Code(9) = 843649688964536.7412@: z_Code(10) = -330085705188364.0817@: z_Code(11) = 41621208.9739@: z_Code(12) = -900372920033759.9903@: z_Code(13) = 291516653989344.1016@: z_Code(14) = -621553923181.6984@: z_Code(15) = 291551690021556.6453@: z_Code(16) = 28798458374890.8543@: z_Code(17) = 86444073845629.4399@: z_Code(18) = 636540268579660.4789@: z_Code(19) = 60911183420250.2143@: z_Code(20) = 846934495644380.8767@: z_Code(21) = 14073829823.4668@: z_Code(22) = 501055845239149.5051@: z_Code(23) = 175724720056981.1236@: z_Code(24) = 75457451135513.7931@: z_Code(25) = -576850389355798.3357@: z_Code(26) = 146298060653075.5445@: z_Code(27) = 850256350680294.7583@: z_Code(28) = -4888724176660.092@: z_Code(29) = 21456079546.6867@
50   zMap VarPtr(z_Code(0))                       'Map z_Code() to z_Data()
60   z_Data(IDX_EBMODE) = zFnAddr("vba6", "EbMode") 'Store the EbMode function address in the thunk data
70   z_Data(IDX_CWP) = zFnAddr("user32", "CallWindowProcA") 'Store CallWindowProc function address in the thunk data
80   z_Data(IDX_SWL) = zFnAddr("user32", "SetWindowLongA") 'Store the SetWindowLong function address in the thunk data
90   z_Data(IDX_FREE) = zFnAddr("kernel32", "VirtualFree") 'Store the VirtualFree function address in the thunk data
100   zMap z_DataOrigData                         'Restore z_Data()'s original data pointer
110   Set z_hWnds = New Collection                'Create the hWnd/thunk-address collection
End Sub
'*****************************************************
Private Sub Class_Terminate()
'*****************************************************
   'Class termination
10   Terminate                                    'Terminate all subclassing
End Sub
'*****************************************************
Public Sub DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As eMsg, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
'*****************************************************
   'Delete the message value from the window handle's specified callback table
    'xTrace "DelMsg - hWnd:" & lng_hWnd & ", uMsg:" & Hex$(uMsg)
10   If When And MSG_BEFORE Then                  'If the message is to be deleted from the before original WndProc table...
20      zDelMsg lng_hWnd, uMsg, IDX_BTABLE        'Delete the message from the before table
30      End If
40   If When And MSG_AFTER Then                   'If the message is to be deleted from the after original WndProc table...
50      zDelMsg lng_hWnd, uMsg, IDX_ATABLE        'Delete the message from the after table
60      End If
70   zMap z_DataOrigData                          'Restore z_Data()'s original data pointer
End Sub
'*****************************************************
Public Function Subclass(ByVal lng_hWnd As Long, ByVal oOwner As iSubclass) As Boolean
'*****************************************************
   'Subclass the specified window handle, calling back to the owners implemented interface
    'xTrace "Subclass - hWnd:" & lng_hWnd
   Dim nAddr               As Long
   Dim nCallback           As Long
10   If IsWindow(lng_hWnd) = 0 Then               'Ensure the window handle is valid
20      zError "Subclass", "wHnd:" & lng_hWnd & " is Invalid"
30      End If
40   RtlMoveMemory VarPtr(nCallback), ObjPtr(oOwner), 4 'Get the address of the owner's vTable
50   zMap nCallback + &H1C                        'Map the implemented interface
60   nCallback = z_Data(0)                        'Get the address of the implemented interface
70   nAddr = VirtualAlloc(0, MEmlEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
80   RtlMoveMemory nAddr, VarPtr(z_Code(0)), CODE_LEN 'Copy the thunk-data/code to the allocated memory
90   On Error GoTo Catch                          'Catch double subclassing
100   z_hWnds.Add nAddr, "h" & lng_hWnd           'Add the hWnd/thunk-address to the collection
110   On Error GoTo 0
120   zMap nAddr                                  'Map z_Data() to the subclass thunk machine-code
130   z_Data(IDX_EBX) = nAddr                     'Patch the data address
140   z_Data(IDX_HWND) = lng_hWnd                 'Store the window handle in the thunk data
150   z_Data(IDX_BTABLE) = nAddr + CODE_LEN       'Store the address of the before table in the thunk data
160   z_Data(IDX_ATABLE) = z_Data(IDX_BTABLE) + ((MSG_ENTRIES + 1) * 4) 'Store the address of the after table in the thunk data
170   z_Data(IDX_OWNER) = ObjPtr(oOwner)          'Store the owner object's address in the thunk data
180   z_Data(IDX_CALLBACK) = nCallback            'Store the zWndProc address in the thunk data
190   nAddr = nAddr + WNDPROC_OFF                 'Execution address of the thunk's WndProc
200   z_Data(IDX_WNDPROC) = SetWindowLongA(lng_hWnd, GWL_WNDPROC, nAddr) 'Set the new WndProc and store the original WndProc in the thunk data
210   zMap z_DataOrigData                         'Restore z_Data()'s original data pointer
220   Subclass = True                             'Indicate success
230   Exit Function                               'Exit
240   Catch:
250   zError "Subclass", "wHnd:" & lng_hWnd & " is already subclassed"
260   On Error GoTo 0
End Function
'*****************************************************
Public Sub Terminate()
'*****************************************************
   'Terminate all subclassing
   Dim i                   As Long
   Dim nAddr               As Long
10   If Not z_hWnds Is Nothing Then               'Ensure that subclassing has been started
        ''''    zError "UnSubAll", "Subclassing hasn't been started", False
        ''''  Else
20      With z_hWnds
30         For i = .Count To 1 Step -1            'Loop through the collection of window handles in reverse order
40            nAddr = .Item(i)                    'Map z_Data() to the hWnd thunk address
50            If IsBadCodePtr(nAddr) = 0 Then     'Ensure that the thunk hasn't already freed itself
60               zMap nAddr                       'Map the thunk memory to the z_Data() array
70               UnSubclass z_Data(IDX_HWND)      'UnSubclass
80               End If
90            Next i                              'Next member of the collection
100         End With
110      Set z_hWnds = Nothing                    'Destroy the hWnd/thunk-address collection
120      End If
End Sub
'*****************************************************
Public Sub UnSubclass(ByVal lng_hWnd As Long)
'*****************************************************
   'UnSubclass the specified window handle
    'xTrace "UnSubclass - hWnd:" & lng_hWnd
10   If z_hWnds Is Nothing Then                   'Ensure that subclassing has been started
20      zError "UnSubclass", "Subclassing hasn't been started", False
30   Else
40      zDelMsg lng_hWnd, eMsg.ALL_MESSAGES, IDX_BTABLE 'Delete all before messages
50      zDelMsg lng_hWnd, eMsg.ALL_MESSAGES, IDX_ATABLE 'Delete all after messages
60      zMap_hWnd lng_hWnd                        'Map the thunk memory to the z_Data() array
70      z_Data(IDX_SHUTDOWN) = -1                 'Set the shutdown indicator
80      zMap z_DataOrigData                       'Restore z_Data()'s original data pointer
90      z_hWnds.Remove "h" & lng_hWnd             'Remove the specified window handle from the collection
100      End If
End Sub
'*****************************************************
Private Sub zAddMsg(ByVal lng_hWnd As Long, ByVal uMsg As eMsg, ByVal nTable As Long)
'*****************************************************
   'Add the message to the specified table of the passed window handle
   Dim i                   As Long               'Loop index
   Dim nCount              As Long               'Table entry count
10   zMap_hWnd lng_hWnd                           'Map z_Data() to the thunk of the specified window handle
20   zMap z_Data(nTable)                          'Map z_Data() to the table address
30   If uMsg = eMsg.ALL_MESSAGES Then             'If ALL_MESSAGES are being added to the table...
40      nCount = eMsg.ALL_MESSAGES                'Set the table entry count to ALL_MESSAGES
50   Else
60      nCount = z_Data(0)                        'Get the current table entry count
70      If nCount >= MSG_ENTRIES Then             'Check for message table overflow
80         zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values", False
90         Exit Sub
100         End If
110      For i = 1 To nCount                      'Loop through the table entries
120         If z_Data(i) = 0 Then                 'If the element is free...
130            z_Data(i) = uMsg                   'Use this element
140            Exit Sub                           'Bail
150         ElseIf z_Data(i) = uMsg Then          'If the message is already in the table...
160            Exit Sub                           'Bail
170            End If
180         Next i                                'Next message table entry
190      nCount = i                               'On drop through: i = nCount + 1, the new table entry count
200      z_Data(nCount) = uMsg                    'Store the message in the appended table entry
210      End If
220   z_Data(0) = nCount                          'Store the new table entry count
End Sub
'*****************************************************
Private Sub zDelMsg(ByVal lng_hWnd As Long, ByVal uMsg As eMsg, ByVal nTable As Long)
'*****************************************************
   'Delete the message from the specified table of the passed window handle
   Dim i                   As Long               'Loop index
   Dim nCount              As Long               'Table entry count
10   zMap_hWnd lng_hWnd                           'Map z_Data() to the thunk of the specified window handle
20   zMap z_Data(nTable)                          'Map z_Data() to the table address
30   If uMsg = eMsg.ALL_MESSAGES Then             'If ALL_MESSAGES are being deleted from the table...
40      z_Data(0) = 0                             'Zero the table entry count
50   Else
60      nCount = z_Data(0)                        'Get the table entry count
70      For i = 1 To nCount                       'Loop through the table entries
80         If z_Data(i) = uMsg Then               'If the message is found...
90            z_Data(i) = 0                       'Null the msg value -- also frees the element for re-use
100            Exit Sub                           'Exit
110            End If
120         Next i                                'Next message table entry
130      zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table", False
140      End If
End Sub
'*****************************************************
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String, Optional ByVal bEnd As Boolean = True)
'*****************************************************
   'Error handler
   'Yorgi - Forgive me Paul, but Msgbox's become a nightmare to stop when cascading events fire, so let's just notify via debug messages!!!!
    ''''App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
    ''''20   MsgBox sMsg & ".", IIf(bEnd, vbCritical, vbExclamation) + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
    ''''  If bEnd Then
    ''''    End
    ''''  End If
    'xTrace TypeName(Me) & "." & sRoutine & ": " & sMsg
#If DebugMode Then
10   Beep
#End If
     'xTrace TypeName(Me) & "." & sRoutine & ": " & sMsg
End Sub
'*****************************************************
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
'*****************************************************
   'Return the address of the specified DLL/procedure
10   zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc) 'Get the specified procedure address
     'debug.Assert zFnAddr                         'In the IDE, validate that the procedure address was located
End Function
'*****************************************************
Private Sub zMap(ByVal nAddr As Long)
'*****************************************************
   'Map the z_Data() array to the specified address
10   RtlMoveMemory z_DataDataPtr, VarPtr(nAddr), 4 'Set z_Data()'s SafeArray data pointer to the specified address
End Sub
'*****************************************************
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
'*****************************************************
   'Map the z_Data() array to the thunk address of the specified window handle
10   If z_hWnds Is Nothing Then                   'Ensure that subclassing has been started
20      zError "zMap_hWnd", "Subclassing hasn't been started", True
30   Else
40      On Error GoTo Catch                       'Catch unsubclassed window handles
50      zMap_hWnd = z_hWnds("h" & lng_hWnd)       'Get the thunk address
60      zMap zMap_hWnd                            'Map z_Data() to the thunk address
70      End If
80   Exit Function                                'Exit returning the thunk address
90   Catch:
100   zError "zMap_hWnd", "Window handle isn't subclassed"
End Function

' Yorgi's 4Matz [Feb 28,2007 23:58:50] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

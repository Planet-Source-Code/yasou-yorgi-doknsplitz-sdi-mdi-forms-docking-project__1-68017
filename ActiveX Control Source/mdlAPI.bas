Attribute VB_Name = "mdlAPI"
Attribute VB_Description = "A module to declare Windows API procedures, functions, types and constants"
'*******************************************************************************
'** File Name   : mdlAPI.bas                                                  **
'** Description : A module to declare Windows API procedures, functions,      **
'**               types and constants                                         **
'*******************************************************************************
Option Explicit
'--- API Functions Parameters Constant
Public Const GA_PARENT                 As Long = &H1 'used by GetAncestor
Public Const GA_ROOT                   As Long = &H2 'used by GetAncestor
Public Const GA_ROOTOWNER              As Long = &H3 'used by GetAncestor
Public Const LB_GETITEMHEIGHT          As Long = &H1A1 'used by SendMessage
' Used by DrawEdge
Public Const BDR_RAISEDOUTER           As Long = &H1
Public Const BDR_SUNKENOUTER           As Long = &H2
Public Const BDR_RAISEDINNER           As Long = &H4
Public Const BDR_SUNKENINNER           As Long = &H8
Public Const EDGE_RAISED               As Long = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN               As Long = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const BF_LEFT                  As Long = &H1
Private Const BF_TOP                   As Long = &H2
Private Const BF_RIGHT                 As Long = &H4
Private Const BF_BOTTOM                As Long = &H8
Public Const BF_RECT                   As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
'--- API Window Messages Parameter
Public Const WA_INACTIVE               As Long = &H0 'parameter of WM_ACTIVATE
'------------------------------------------
' Refer to MSDN Library for the complete
' descripion of each API declaration below
'------------------------------------------
'--- API Function/Sub Declaration
' Purpose    - Converts the client coordinates of a specified point to screen
'              coordinates
Public Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI)
' Purpose    - Confines the cursor to a rectangular area lpRect on the screen
Public Declare Sub ClipCursor Lib "user32" (lpRect As RECT)
Attribute ClipCursor.VB_Description = "Confines the cursor to a rectangular area lpRect on the screen"
' Purpose    - Frees the cursor to move anywhere on the screen
Public Declare Sub ClipCursorClear Lib "user32" Alias "ClipCursor" (Optional ByVal lpRect As Long = 0&)
Attribute ClipCursorClear.VB_Description = "Frees the cursor to move anywhere on the screen"
' Purpose    - Retrieves the handle to the ancestor of the specified window
Public Declare Function GetAncestor Lib "user32" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
' Purpose    - Retrieves the handle to the window (if any) that has captured the
'              mouse or stylus input
' Usage      : In ctlButton to provide hover effect along with SetCapture,
'              ReleaseCapture and WindowFromPoint
Public Declare Function GetCapture Lib "user32" () As Long
' Purpose    - Retrieves the cursor's position in screen coordinates
Public Declare Sub GetCursorPos Lib "user32" (lpPoint As POINTAPI)
Attribute GetCursorPos.VB_Description = "Retrieves the cursor's position in screen coordinates"
' Purpose    - Releases the mouse capture from a window in the current thread
'              and restores normal mouse input processing
' Usage      : In ctlButton to provide hover effect along with SetCapture,
'              GetCapture and WindowFormPoint
Public Declare Function ReleaseCapture Lib "user32" () As Long
' Purpose    - Converts the screen coordinates of a specified point on the
'              screen to client coordinates
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
' Purpose    - Sends message wMsg to window hWnd
' Usage      : Gets the height of the item in list box control or other controls
'              that inherit it
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Attribute SendMessage.VB_Description = "Sends message wMsg to window hWnd"
' Purpose    - Sets the mouse capture to the specified window belonging to the
'              current thread
' Usage      : In ctlButton to provide hover effect along with GetCapture,
'              ReleaseCapture and WindowFromPoint
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
' Purpose    - Sets the coordinates of the specified rectangle
' Usage      : Just to shorten the number of source code lines needed to set
'              the rectangle coordinate
Public Declare Sub SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
' Purpose    - Performs an operation on a specified file
' Usage      : Opens mail client for specified e-mail address
Public Declare Sub ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
' Purpose    - Retrieves the handle to the window that contains the specified
'              point
' Usage      : In ctlButton to provide hover effect along with SetCapture,
'              GetCapture and ReleaseCapture
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
''''Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Any) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hChildwnd As Long, ByVal hParentwnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetPropA Lib "user32" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetPropA Lib "user32" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemovePropA Lib "user32" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub CopyMemoryRect Lib "kernel32" Alias "RtlMoveMemory" (Destination As RECT, ByVal Source As Long, ByVal Length As Long)
Public Declare Sub CopyMemoryFromRect Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As RECT, ByVal Length As Long)
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long

' Yorgi's 4Matz [Feb 28,2007 23:58:49] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

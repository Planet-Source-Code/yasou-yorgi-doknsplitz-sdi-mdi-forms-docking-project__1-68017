Attribute VB_Name = "modDefsWin32Msgs"
Option Explicit
Public Const EM_GETLIMITTEXT           As Long = &HD5
Public Const EM_LIMITTEXT              As Long = &HC5
Public Const EM_GETSEL                 As Long = &HB0
Public Const SMTO_NORMAL               As Long = &H0
Public Const SC_RESTORE                As Long = &HF120&
Public Const SC_CLOSE                  As Long = &HF060&
Public Const SC_MOVE                   As Long = &HF010&
Public Const SC_SIZE                   As Long = &HF000&
'''''Get Window ------------------------
Public Const GW_CHILD                  As Long = 5
Public Const GW_HWNDNEXT               As Integer = 2
Public Const GW_HWNDPREV               As Integer = 3
Public Const GWL_STYLE                 As Long = (-16)
Public Const GWL_EXSTYLE               As Long = (-20)
Public Const GWL_HWNDPARENT            As Long = (-8)
'''''Window Styles------------------------
Public Const WS_EX_WINDOWEDGE          As Long = &H100
Public Const WS_EX_CLIENTEDGE          As Long = &H200
Public Const WS_EX_STATICEDGE          As Long = &H20000
Public Const WS_EX_LAYERED             As Long = &H80000
''''Public Const WS_MINIMIZEBOX                         As Long = &H20000
''''Public Const WS_MAXIMIZEBOX                         As Long = &H10000
''''Public Const WS_VSCROLL                             As Long = &H200000
''''Public Const WS_VISIBLE                             As Long = &H10000000
''''Public Const WS_CHILD                               As Long = &H40000000
Public Const WS_BORDER                 As Long = &H800000
Public Const WS_CAPTION                As Long = &HC00000
Public Const WS_CAPTION_NOT            As Long = &HFFFFFFFF - WS_CAPTION
Public Const WS_ACTIVECAPTION          As Long = &H1
Public Const WS_CHILD                  As Long = &H40000000
Public Const WS_CHILDWINDOW            As Long = (WS_CHILD)
Public Const WS_CLIPCHILDREN           As Long = &H2000000
Public Const WS_CLIPSIBLINGS           As Long = &H4000000
Public Const WS_DISABLED               As Long = &H8000000
Public Const WS_DLGFRAME               As Long = &H400000
Public Const WS_GROUP                  As Long = &H20000
Public Const WS_TABSTOP                As Long = &H10000
Public Const WS_GT                     As Double = WS_GROUP Or WS_TABSTOP
Public Const WS_HSCROLL                As Long = &H100000
Public Const WS_MAXIMIZE               As Long = &H1000000
Public Const WS_MINIMIZE               As Long = &H20000000
Public Const WS_ICONIC                 As Long = WS_MINIMIZE
Public Const WS_MAXIMIZEBOX            As Long = &H10000
Public Const WS_MINIMIZEBOX            As Long = &H20000
Public Const WS_OVERLAPPED             As Long = &H0&
Public Const WS_SYSMENU                As Long = &H80000
Public Const WS_THICKFRAME             As Long = &H40000
Public Const WS_OVERLAPPEDWINDOW       As Double = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
Public Const WS_POPUP                  As Long = &H80000000
Public Const WS_POPUPWINDOW            As Double = WS_POPUP Or WS_BORDER Or WS_SYSMENU
Public Const WS_SIZEBOX                As Long = WS_THICKFRAME
Public Const WS_TILED                  As Long = WS_OVERLAPPED
Public Const WS_TILEDWINDOW            As Long = WS_OVERLAPPEDWINDOW
Public Const WS_VISIBLE                As Long = &H10000000
Public Const WS_VSCROLL                As Long = &H200000
Public Const WS_EX_TOOLWINDOW          As Long = &H80&
Public Const WS_EX_APPWINDOW           As Long = &H40000
Public Const SWP_NOOWNERZORDER         As Long = &H200     ' Don"t do owner Z ordering
Public Const SWP_FRAMECHANGED          As Long = &H20
Public Const SWP_NOREPOSITION          As Long = SWP_NOOWNERZORDER
Public Const SWP_NOZORDER              As Long = &H4
Public Const SWP_NOACTIVATE            As Long = &H10
Public Const SWP_SHOWWINDOW            As Long = &H40
Public Const SWP_NOSIZE                As Long = &H1
Public Const SWP_NOMOVE                As Long = &H2
Public Const TOPMOST_FLAGS             As Double = SWP_NOMOVE Or SWP_NOSIZE
Public Const SW_SHOW                   As Integer = 5
Public Const SW_HIDE                   As Integer = 0
Public Const SW_SHOWNORMAL             As Integer = 1
Public Const SW_MAXIMIZE               As Integer = 3
Public Const MIIM_STATE                As Long = &H1&
Public Const MIIM_ID                   As Long = &H2&
Public Const MFS_GRAYED                As Long = &H3&
Public Const SIZE_MINIMIZED            As Long = &H1&
Public Type MENUITEMINFO
   cbSize                              As Long
   fMask                               As Long
   fType                               As Long
   fState                              As Long
   wID                                 As Long
   hSubMenu                            As Long
   hbmpChecked                         As Long
   hbmpUnchecked                       As Long
   dwItemData                          As Long
   dwTypeData                          As String
   cch                                 As Long
End Type
Public Enum OMNIWINDOWSMSG
   WM_USER_TEXT_CHANGED = &H400 + &H1
   WM_USER_SB_GETRECT = &H400 + 10&
End Enum
Public Type COPYDATASTRUCT
   dwData                              As Long
   cbData                              As Long
   lpData                              As Long
End Type

' Yorgi's 4Matz [Feb 28,2007 23:58:49] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

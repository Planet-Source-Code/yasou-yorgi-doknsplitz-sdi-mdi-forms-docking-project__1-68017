VERSION 1.0 CLASS
BEGIN
MultiUse = -1                                     'True
Persistable = 0                                   'NotPersistable
DataBindingBehavior = 0                           'vbNone
DataSourceBehavior  = 0                           'vbNone
MTSTransactionMode  = 0                           'NotAnMTSObject
END
Attribute VB_Name = "iSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*************************************************************************************************
'* iSubclass - cSubclass' implementation interface class + enumerations
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 The original..................................................................... 20060322
'* 01/21/07 TODO Yorgi: Chg wParam & lParam to BYREF. Some API calls (ie CopyMemoryFromRect) require data to be passed back!
'*************************************************************************************************
Option Explicit
'-cSublass public declarations--------------------------------------------------------------------
Public Enum eMsgWhen                              'When to callback
   MSG_BEFORE = 1                                 'Callback before the original WndProc
   MSG_AFTER = 2                                  'Callback after the original WndProc
   MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER     'Callback before and after the original WndProc
End Enum
'*****************************************************
Public Sub WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As eMsg, ByVal wParam As Long, ByVal lParam As Long)
'*****************************************************
   '* bBefore  - Indicates whether the callback is before or after the original WndProc. Usually you
   '*            will know unless the callback for the uMsg value is specified as MSG_BEFORE_AFTER
   '*            (both before and after the original WndProc).
   '* bHandled - In a before original WndProc callback, setting bHandled to True will prevent the
   '*            message being passed to the original WndProc and (if set to do so) the after
   '*            original WndProc callback.
   '* Return   - WndProc return value. Set as per the MSDN documentation for the message value,
   '*            and/or, in an after the original WndProc callback, act on the return value set by
   '*            the original WndProc.
   '* hWnd     - Window handle.
   '* uMsg     - Message value.
   '* wParam   - Message related data.
   '* lParam   - Message related data.
   '*************************************************************************************************
End Sub

' Yorgi's 4Matz [Feb 28,2007 23:58:50] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

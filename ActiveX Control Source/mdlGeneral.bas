Attribute VB_Name = "mdlGeneral"
Attribute VB_Description = "A module to handle general operations"
'*******************************************************************************
'** File Name   : mdlGeneral.bas                                              **
'** Description : A module to handle general operations                       **
'*******************************************************************************
Option Explicit
'--- Resource File Constants
' Splitter Cursor
Public Const gconCurHSplitter          As Long = 101 'horizontal splitter cursor
Public Const gconCurVSplitter          As Long = 102 'vertical splitter cursor
' Error Message Index
Public Enum genmErrNumber
   errBuildSplitters = 2000
   errSelfContained = 2001
   errMoveSplitter = 2002
   errResizeSplitter = 2003
   errMoveControlRoom = 2004
   errIdControl = 2005
   errIdSplitter = 2006
   errMoveControlClosed = 2007
End Enum
'--- Other Constants
Public Const gconUninitializedLong     As Long = -1 'represent the Id which is not exist or hasn't been initialized yet
Public Const gconLngInfinite           As Long = 2147483647
Public Const gconPROPERTY_DFPTR        As String = "dokfrm" 'Win GetProp/SetProp property name
Public Const gconSettingDocking        As String = "Docking" 'VB Get/Set Settings key name
'--- Variable Declaration
Public gstrControlName                 As String  'the name of DoknSplitz
#If DebugMode Then
Public gstrParentName                  As String  'the name of Parent DoknSplitz
Private lCounter                       As Long
#End If
'*****************************************************
Sub AppTrace(sOwner As String, sProc As String, sMsg As String) '4Matz:New
'*****************************************************
   Dim lHandle             As Long
   Dim sBuffer             As String
#If DebugMode Then
10   If lCounter < 9999 Then
20      lCounter = lCounter + 1
30   Else
40      lCounter = 1
50      End If
60   sBuffer = Format$(lCounter, "0000 ") & sOwner & "(" & sProc & ")." & sMsg
#If DebugMode = 2 Then
70   lHandle = FreeFile
80   Open "c:\temp\TraceLog.txt" For Append As #lHandle
90   Print #lHandle, sBuffer
100   Close #lHandle
#Else
110   Debug.Print sBuffer
#End If
#End If
End Sub
'*****************************************************
Public Sub ErrHandler(oError As ErrObject, sErrStr As String, strProc As String, strModule As String) '4Matz:New
'*****************************************************
   Dim lngErr              As Long
   Dim MsgBoxStr           As String
   Dim strError            As String
   Dim strTitle            As String
10   lngErr = oError.Number
20   strTitle = oError.Source
30   If lngErr > 1000 Then
40      If LenB(sErrStr) = 0 Then
50         sErrStr = LoadResString(lngErr)
60         End If
70      End If
80   MsgBoxStr = "Error....: " & sErrStr & vbNewLine
90   MsgBoxStr = MsgBoxStr & "ErrorNo..: " & lngErr & vbNewLine
100   MsgBoxStr = MsgBoxStr & "Module...: " & gstrControlName & ":" & strModule & vbNewLine
110   MsgBoxStr = MsgBoxStr & "Procedure: " & strProc & vbNewLine
      '  MsgBoxStr = MsgBoxStr & "Line No..: " & ErrLine
120   Beep
130   If LenB(strTitle) = 0 Then
140      strTitle = "Application Error"
150      End If
160   MsgBox MsgBoxStr, vbCritical, strTitle
170   AppTrace strModule, strProc, Replace$(MsgBoxStr, vbNewLine, "|")
#If DebugMode = 3 Then
180   Stop                                        'stop here, F8 moves to error line
#End If
End Sub
'*****************************************************
Public Function GetCursorRelPos(hWnd As Long, Optional lCurrX As Long, Optional lCurrY As Long) As POINTAPI '4Matz:Changed
'*****************************************************
   ' Purpose    - Retrieves the cursor's position in twips relative to certain window 
   ' Assumptions: Window hwnd exist (if hwnd is not omitted) 
   ' Input      - hwnd (the window where the cursor will be retrieved relative to; 
   '                    if ommited, the screen will be used as the window) 
   ' Return     : As specified 
   Dim uposGetCursorRelPos As POINTAPI
10   GetCursorPos uposGetCursorRelPos
20   If Not IsMissing(lCurrX) Then
30      lCurrX = uposGetCursorRelPos.X
40      lCurrY = uposGetCursorRelPos.Y
50      End If
60   If hWnd <> gconUninitializedLong Then
70      ScreenToClient hWnd, uposGetCursorRelPos
80      With uposGetCursorRelPos
90         .X = .X * Screen.TwipsPerPixelX
100         .Y = .Y * Screen.TwipsPerPixelY
110         End With
120      End If
130   GetCursorRelPos = uposGetCursorRelPos
End Function
'*****************************************************
Public Function GetMin(ParamArray vntValue() As Variant) As Long  '4Matz:Changed
Attribute GetMin.VB_Description = "Gets minimum value of numbers in array lngValue()"
'*****************************************************
   ' Purpose    - Gets minimum value of numbers in array lngValue()
   ' Assumptions: * Option base is set to 0
   '              * Array lngValue() contains only numbers
   ' Input      - vntValue()
   ' Return     : * If no parameters passed to vntValue(), returns Empty
   '              * Otherwise, returns as specified
   '** 01/26/07 Yorgi- Performance & cleanup, redef variants to longs
   Dim i                   As Long               'for iterating the parameters value
   Dim lVal                As Long               'returned value
   Dim vntGetMin           As Long               'returned value
10   If Not IsMissing(vntValue) Then
20      vntGetMin = CLng(vntValue(0))
30      For i = 1 To UBound(vntValue)
40         lVal = CLng(vntValue(i))
50         If lVal < vntGetMin Then vntGetMin = lVal
60         Next
70      GetMin = vntGetMin
80      End If
End Function
'*****************************************************
Public Function HiWord(lDWord As Long) As Integer '4Matz:New
'*****************************************************
10   HiWord = (lDWord And &HFFFF0000) \ &H10000
End Function
'*****************************************************
Public Function LoWord(lDWord As Long) As Integer '4Matz:New
'*****************************************************
10   If lDWord And &H8000& Then
20      LoWord = lDWord Or &HFFFF0000
30   Else
40      LoWord = lDWord And &HFFFF&
50      End If
End Function
'*****************************************************
Public Property Get ObjectFromPtr(ByVal lPtr As Long) As Object '4Matz:New
'*****************************************************
   Dim objT                As Object
10   If Not (lPtr = 0) Then
20      CopyMemory objT, lPtr, 4
30      Set ObjectFromPtr = objT
40      CopyMemory objT, 0&, 4
50      End If
End Property
'*****************************************************
Public Sub RaiseError(ByVal udeErrNumber As genmErrNumber, Optional ByVal strSource As String)
Attribute RaiseError.VB_Description = "Raises custom error udeErrNumber"
'*****************************************************
   ' Purpose    - Raises custom error udeErrNumber
   ' Assumptions: * Error message udeErrNumber exists in the resource file
   '              * Global variable gstrControlName has been initialized
   ' Inputs     - * udeErrNumber
   '              * strSource (the location in form ClassNaRoutinesName where
   '                the error occur
10   If strSource <> "." Then strSource = "." & strSource
20   Err.Raise (vbObjectError + udeErrNumber), gstrControlName & strSource, LoadResString(udeErrNumber)
End Sub

' Yorgi's 4Matz [Feb 28,2007 23:58:49] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

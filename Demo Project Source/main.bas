Attribute VB_Name = "modMain"
Option Explicit
Private Const MODULENAME               As String = "modMain"
Public Const mconNameDemoFeatures      As String = "frmDemoFeatures"
Public Const mconNameDemoEnhancements  As String = "frmDemoEnhancements"
Public Const mconNameDonaldDuck        As String = "frmDonaldDuck"
Public Const mconNameMickeyMouse       As String = "frmMickeyMouse"
Public Const mconNameLabel1            As String = "Label1"
'*****************************************************
Public Sub ArrangeControls(oHostFrm As Form)
'*****************************************************
   'stack Mickey & Donald around the Label1 control
   Dim oVCtlLabel          As clsControl
10   With oHostFrm.ControlManager1
20      Set oVCtlLabel = .Controls(mconNameLabel1)
        'Arrangement depends on position of Label1 control so make sure its active
30      If Not oVCtlLabel.Closed Then
40         If oHostFrm.mnuMickysForum.Enabled = False Then
50            If oHostFrm.mnuDonaldsForum.Enabled = False Then
60               If .MoveControl(.Controls(.DockedForm(frmMickeyMouse).VCtlKey), mdControlTop, mconNameLabel1) Then
70                  If .MoveControl(.Controls(.DockedForm(frmDonaldDuck).VCtlKey), mdControlBottom, mconNameLabel1) Then
80                     If .MoveSplitter(oVCtlLabel.IdSplRight, .Width * 0.6) Then
90                        If .MoveSplitter(oVCtlLabel.IdSplLeft, .Width * 0.4) Then
100                           If .MoveSplitter(oVCtlLabel.IdSplBottom, .Height * 0.5) Then
110                              End If
120                           End If
130                        End If
140                     End If
150                  End If
160               End If
170            End If
180         End If
190      End With
End Sub
'*****************************************************
Public Sub ErrHandler(oError As ErrObject, sErrStr As String, strProc As String, strModule As String)
'*****************************************************
   Dim lngErr              As Long
   Dim MsgBoxStr           As String
   Dim sTitle              As String
   Dim strError            As String
10   lngErr = oError.Number
20   sTitle = oError.Source
30   If lngErr > 1000 Then
40      If LenB(sErrStr) = 0 Then
50         sErrStr = LoadResString(lngErr)
60         End If
70      End If
80   MsgBoxStr = "Error....: " & sErrStr & vbNewLine
90   MsgBoxStr = MsgBoxStr & "ErrorNo..: " & lngErr & vbNewLine
100   MsgBoxStr = MsgBoxStr & "Module...: " & ":" & strModule & vbNewLine
110   MsgBoxStr = MsgBoxStr & "Procedure: " & strProc & vbNewLine
120   Beep
130   If LenB(sTitle) = 0 Then
140      sTitle = "Application Error"
150      End If
160   MsgBox MsgBoxStr, vbCritical, sTitle
      'debug.print Replace$(MsgBoxStr, vbNewLine, "|")
End Sub
'*****************************************************
Sub main()
'*****************************************************
10   frmSDIMain.Show
End Sub
'*****************************************************
Public Sub RebuildDemo(oHostFrm As Form)
'*****************************************************
   'FormAdd will create new DoknForm objects if they do not exist, otherwise
   ' we Open the existing object and positions are not changed!
   'If you want to force a particular position, first make sure object is not .Closed,
   '  then use the MoveControl function for specific placement
   Dim oHostDS             As DoknSplitz.ControlManager
10   Set oHostDS = oHostFrm.ControlManager1
     'order of FormAdd's does effect initial position placement!
20   oHostDS.BatchBuild = True                    'batch the following FormAdd's for later processing
30   ShowDF oHostDS, frmDemoFeatures
40   ShowDF oHostDS, frmDemoEnhancements
50   ShowDF oHostDS, frmDonaldDuck
60   ShowDF oHostDS, frmMickeyMouse
70   ShowDF oHostDS, oHostFrm.Label1
80   oHostDS.BatchBuild = False                   'reBuild and draw all controls
90   RebuildDemo_Exit:
100   On Error GoTo 0
110   Exit Sub
120   RebuildDemo_Err:
130   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", RebuildDemo", MODULENAME
140   Resume RebuildDemo_Exit
End Sub
'*****************************************************
Public Sub ShowDF(oHostDS As DoknSplitz.ControlManager, oFrmOrCtl As Object)
'*****************************************************
10   With oHostDS
20      Select Case oFrmOrCtl.Name
           Case mconNameDemoFeatures: .FormAdd oFrmOrCtl, , , DAlignRight, , DSLeft Or DSRight
30         Case mconNameDemoEnhancements: .FormAdd oFrmOrCtl, , , DAlignLeft
40         Case mconNameDonaldDuck: .FormAdd oFrmOrCtl, , , DAlignLeft, , , TBO_VERTICAL
50         Case mconNameMickeyMouse: .FormAdd oFrmOrCtl, , , DAlignLeft, , DSTop Or DSLeft Or DSRight Or DSFloat, TBO_VERTICAL
60         Case mconNameLabel1: .ShowControl mconNameLabel1, True 'show the Label1 design time control
70         End Select
80      End With
End Sub

' Yorgi's 4Matz [Feb 28,2007 23:58:48] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

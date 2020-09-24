VERSION 5.00
Object = "{5906E796-EE78-4E1C-BEE0-327463DEA5CC}#54.0#0"; "DokNSplitz.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmDemoEnhancements
Caption         =   "Yorgi's Enhancements"
ClientHeight    =   6240
ClientLeft      =   60
ClientTop       =   450
ClientWidth     =   6660
LinkTopic       =   "Form1"
ScaleHeight     =   6240
ScaleWidth      =   6660
Visible         =   0                             'False
Begin DoknSplitz.ControlManager ControlManager1
Height          =   5895
Left            =   150
TabIndex        =   0
Top             =   120
Width           =   5925
_ExtentX        =   10451
_ExtentY        =   10398
LiveUpdate      =   0                             'False
TitleBar_TBarType=   1
Begin SHDocVwCtl.WebBrowser WBEvents
Height          =   2543
Left            =   3109
TabIndex        =   3
Top             =   3352
Width           =   2816
ExtentX         =   4967
ExtentY         =   4486
ViewMode        =   0
Offline         =   0
Silent          =   0
RegisterAsBrowser=   0
RegisterAsDropTarget=   1
AutoArrange     =   0                             'False
NoClientEdge    =   0                             'False
AlignLeft       =   0                             'False
NoWebView       =   0                             'False
HideFileNames   =   0                             'False
SingleClick     =   0                             'False
SingleSelection =   0                             'False
NoFolders       =   0                             'False
Transparent     =   0                             'False
ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
Location        =   "http:///"
End
Begin SHDocVwCtl.WebBrowser WBMethods
Height          =   2842
Left            =   3109
TabIndex        =   2
Top             =   225
Width           =   2816
ExtentX         =   4967
ExtentY         =   5013
ViewMode        =   0
Offline         =   0
Silent          =   0
RegisterAsBrowser=   0
RegisterAsDropTarget=   1
AutoArrange     =   0                             'False
NoClientEdge    =   0                             'False
AlignLeft       =   0                             'False
NoWebView       =   0                             'False
HideFileNames   =   0                             'False
SingleClick     =   0                             'False
SingleSelection =   0                             'False
NoFolders       =   0                             'False
Transparent     =   0                             'False
ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
Location        =   "http:///"
End
Begin SHDocVwCtl.WebBrowser WBIntro
Height          =   5670
Left            =   0
TabIndex        =   1
Top             =   225
Width           =   3049
ExtentX         =   5378
ExtentY         =   10001
ViewMode        =   0
Offline         =   0
Silent          =   0
RegisterAsBrowser=   0
RegisterAsDropTarget=   1
AutoArrange     =   0                             'False
NoClientEdge    =   0                             'False
AlignLeft       =   0                             'False
NoWebView       =   0                             'False
HideFileNames   =   0                             'False
SingleClick     =   0                             'False
SingleSelection =   0                             'False
NoFolders       =   0                             'False
Transparent     =   0                             'False
ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
Location        =   "http:///"
End
End
End
Attribute VB_Name = "frmDemoEnhancements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************
Private Sub ControlManager1_TitleBarCaption(ByVal sIdControl As String, sCaption As String)
'*****************************************************
10   Select Case sIdControl
        Case "WBIntro": sIdControl = "Introduction"
20      Case "WBMethods": sIdControl = "Procedures"
30      Case "WBEvents": sIdControl = "Events"
40      End Select
50   sCaption = "[" & sIdControl & "]"
End Sub
'*****************************************************
Private Sub Form_Load()
'*****************************************************
10   On Error Resume Next                         'incase docs were moved/deleted
20   WBIntro.Navigate App.Path & "\docs\DoknSplitz.htm"
30   WBMethods.Navigate App.Path & "\docs\ctlControlManager.ctl.Procs.4Matz.htm"
40   WBEvents.Navigate App.Path & "\docs\ctlControlManager.ctl.Events.4Matz.htm"
50   DoEvents                                     'MUST GIVE TIME TO COMPLETE TO AVOID IE7 Runtime Error 7
End Sub

' Yorgi's 4Matz [Feb 28,2007 23:58:47] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

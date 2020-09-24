VERSION 5.00
Object = "{5906E796-EE78-4E1C-BEE0-327463DEA5CC}#56.0#0"; "DokNSplitz.ocx"
Begin VB.Form frmSDIMain
BackColor       =   &H8000000C&
Caption         =   "DoknSplitz SDI Demo"
ClientHeight    =   5430
ClientLeft      =   165
ClientTop       =   780
ClientWidth     =   10140
LinkTopic       =   "Form1"
ScaleHeight     =   5430
ScaleWidth      =   10140
StartUpPosition =   2                             'CenterScreen
Visible         =   0                             'False
Begin DoknSplitz.ControlManager ControlManager1
Height          =   5190
Left            =   150
TabIndex        =   0
Top             =   90
Width           =   9765
_ExtentX        =   17224
_ExtentY        =   9155
BackColor       =   -2147483648
FillContainer   =   0                             'False
LiveUpdate      =   0                             'False
TitleBar_TBarType=   1
UnloadFrmOnClose=   -1                            'True
Begin VB.Label Label1
BackColor       =   &H00FFFFFF&
Caption         =   "Design-Time controls get to play too!"
BeginProperty Font
Name            =   "Times New Roman"
Size            =   15.75
Charset         =   0
Weight          =   700
Underline       =   0                             'False
Italic          =   -1                            'True
Strikethrough   =   0                             'False
EndProperty
ForeColor       =   &H000000FF&
Height          =   4965
Left            =   0
TabIndex        =   1
Top             =   225
Width           =   9765
End
End
Begin VB.Menu mnuDemo
Caption         =   "&Demo"
Begin VB.Menu mnuMDI
Caption         =   "Start MDI Demo"
End
Begin VB.Menu mnuRebuild
Caption         =   "&Rebuild"
End
Begin VB.Menu mnuArrangeControls
Caption         =   "Arrange Controls"
End
Begin VB.Menu mnuSeparator
Caption         =   "-"
End
Begin VB.Menu mnuExit
Caption         =   "E&xit"
End
End
Begin VB.Menu mnuEnhancements
Caption         =   "&Enhancements"
Enabled         =   0                             'False
End
Begin VB.Menu mnuFeatures
Caption         =   "&Features"
Enabled         =   0                             'False
End
Begin VB.Menu mnuMickysForum
Caption         =   "Micky's Forum"
Enabled         =   0                             'False
End
Begin VB.Menu mnuDonaldsForum
Caption         =   "Donald's Forum"
Enabled         =   0                             'False
End
Begin VB.Menu mnuAbout
Caption         =   "About"
End
End
Attribute VB_Name = "frmSDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
Private Sub ControlManager1_FormAddComplete(ByVal sFormName As String, ByVal sKey As String, ByVal sIdControl As String)
'*****************************************************
10   UpdateMenu sFormName, False
End Sub
'*****************************************************
Private Sub ControlManager1_FormRemoveComplete(ByVal sFormName As String)
'*****************************************************
10   UpdateMenu sFormName, True
End Sub
'*****************************************************
Private Sub Form_Load()
'*****************************************************
10   frmWait.Show
     ' These Move functions assume startup locations, so valid during demo startup
20   With ControlManager1
30      .LiveUpdate = False                       'faster only to draw the splitter without child control overhead
40      .UnloadFrmOnClose = True                  'click control close button will Unload the form as well
50      .FillContainer = True
60      .TitleBar_TBarType = TBT_DEFAULT
70      End With
80   RebuildDemo Me
90   ArrangeControls Me
100   Unload frmWait
110   DoEvents
End Sub
'*****************************************************
Private Sub Form_Unload(Cancel As Integer)
'*****************************************************
   Dim vVar                As Form
10   ControlManager1.DetachAll                    'not a bad idea to stop all subclassing before unloading forms
20   For Each vVar In Forms
30      If vVar.hWnd <> hWnd Then
40         Unload vVar
50         End If
60      Next
End Sub
'*****************************************************
Private Sub mnuAbout_Click()
'*****************************************************
10   ControlManager1.ShowAboutBox
End Sub
'*****************************************************
Private Sub mnuArrangeControls_Click()
'*****************************************************
10   ArrangeControls Me
End Sub
'*****************************************************
Private Sub mnuDonaldsForum_Click()
'*****************************************************
10   ShowDF ControlManager1, frmDonaldDuck
End Sub
'*****************************************************
Private Sub mnuEnhancements_Click()
'*****************************************************
10   ShowDF ControlManager1, frmDemoEnhancements
End Sub
'*****************************************************
Private Sub mnuExit_Click()
'*****************************************************
10   Unload Me
End Sub
'*****************************************************
Private Sub mnuFeatures_Click()
'*****************************************************
10   ShowDF ControlManager1, frmDemoFeatures
End Sub
'*****************************************************
Private Sub mnuMDI_Click()
'*****************************************************
10   Unload Me
20   DoEvents
30   frmMDIMain.Show
End Sub
'*****************************************************
Private Sub mnuMickysForum_Click()
'*****************************************************
10   ShowDF ControlManager1, frmMickeyMouse
End Sub
'*****************************************************
Private Sub mnuRebuild_Click()
'*****************************************************
10   RebuildDemo Me
20   ArrangeControls Me
End Sub
'*****************************************************
Private Sub UpdateMenu(sFormName As String, bEnable As Boolean)
'*****************************************************
10   Select Case sFormName
        Case mconNameDemoFeatures: mnuFeatures.Enabled = bEnable
20      Case mconNameDemoEnhancements: mnuEnhancements.Enabled = bEnable
30      Case mconNameDonaldDuck: mnuDonaldsForum.Enabled = bEnable
40      Case mconNameMickeyMouse: mnuMickysForum.Enabled = bEnable
50      End Select
End Sub

' Yorgi's 4Matz [Feb 28,2007 23:58:48] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

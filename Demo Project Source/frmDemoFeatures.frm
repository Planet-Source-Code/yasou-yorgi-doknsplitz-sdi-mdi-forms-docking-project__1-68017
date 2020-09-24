VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5906E796-EE78-4E1C-BEE0-327463DEA5CC}#55.0#0"; "DokNSplitz.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmDemoFeatures
Caption         =   "Theo's Features"
ClientHeight    =   9090
ClientLeft      =   165
ClientTop       =   -1350
ClientWidth     =   9300
LinkTopic       =   "Form1"
ScaleHeight     =   9090
ScaleWidth      =   9300
Visible         =   0                             'False
Begin DoknSplitz.ControlManager ControlManager1
Height          =   9090
Left            =   30
TabIndex        =   0
Top             =   30
Width           =   8235
_ExtentX        =   14526
_ExtentY        =   16034
LiveUpdate      =   0                             'False
TitleBar_TBarType=   3
Begin SHDocVwCtl.WebBrowser WBFeatures
Height          =   3105
Left            =   3600
TabIndex        =   194
Top             =   330
Width           =   4575
ExtentX         =   8070
ExtentY         =   5477
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
Location        =   ""
End
Begin VB.PictureBox picMain
Height          =   7462
Left            =   -30
ScaleHeight     =   7395
ScaleWidth      =   3465
TabIndex        =   4
Top             =   300
Width           =   3525
Begin VB.Frame fraFeatures
Height          =   7695
Index           =   6
Left            =   330
TabIndex        =   91
Top             =   390
Width           =   3030
Begin VB.VScrollBar vsbSplProperties
Height          =   7545
Left            =   2805
TabIndex        =   149
Top             =   105
Width           =   210
End
Begin VB.Frame fraConSplProperties
BorderStyle     =   0                             'None
Caption         =   "Frame1"
Height          =   6075
Left            =   120
TabIndex        =   92
Top             =   870
Width           =   2730
Begin VB.Frame fraSplProperties
BorderStyle     =   0                             'None
Caption         =   "Frame1"
Height          =   11595
Left            =   1560
TabIndex        =   93
Top             =   4950
Width           =   2700
Begin VB.ComboBox cboSplOrientation
Enabled         =   0                             'False
Height          =   315
ItemData        =   "frmDemoFeatures.frx":0000
Left            =   1605
List            =   "frmDemoFeatures.frx":000A
Style           =   2                             'Dropdown List
TabIndex        =   118
ToolTipText     =   "Returns the virtual splitter movement direction"
Top             =   11370
Width           =   915
End
Begin VB.TextBox txtSplYc
Enabled         =   0                             'False
Height          =   285
Left            =   1605
TabIndex        =   117
ToolTipText     =   "Returns the y-coordinate of the virtual splitter center"
Top             =   14070
Width           =   915
End
Begin VB.TextBox txtSplXc
Enabled         =   0                             'False
Height          =   285
Left            =   1605
TabIndex        =   116
ToolTipText     =   "Returns the x-coordinate of the virtual splitter center"
Top             =   13530
Width           =   915
End
Begin VB.TextBox txtSplWidth
Enabled         =   0                             'False
Height          =   285
Left            =   1605
TabIndex        =   115
ToolTipText     =   "Returns the width of the virtual splitter"
Top             =   13005
Width           =   915
End
Begin VB.TextBox txtSplTop
Enabled         =   0                             'False
Height          =   285
Left            =   1605
TabIndex        =   114
ToolTipText     =   "Returns the distance between the internal top edge of the virtual splitter and the top edge of the related Control Manager object"
Top             =   12465
Width           =   915
End
Begin VB.TextBox txtSplRight
Enabled         =   0                             'False
Height          =   285
Left            =   1605
TabIndex        =   113
ToolTipText     =   $"frmDemoFeatures.frx":0028
Top             =   11940
Width           =   915
End
Begin VB.TextBox txtSplMinYc
Enabled         =   0                             'False
Height          =   285
Left            =   1605
TabIndex        =   112
ToolTipText     =   "Returns the minimum y-coordinate of the virtual splitter"
Top             =   10845
Width           =   915
End
Begin VB.TextBox txtSplMinXc
Enabled         =   0                             'False
Height          =   285
Left            =   1605
TabIndex        =   111
ToolTipText     =   "Returns the minimum x-coordinate of the virtual splitter"
Top             =   10305
Width           =   915
End
Begin VB.TextBox txtSplMaxYc
Enabled         =   0                             'False
Height          =   285
Left            =   1605
TabIndex        =   110
ToolTipText     =   "Returns the maximum y-coordinate of the virtual splitter"
Top             =   9780
Width           =   915
End
Begin VB.TextBox txtSplMaxXc
Enabled         =   0                             'False
Height          =   285
Left            =   1605
TabIndex        =   109
ToolTipText     =   "Returns the maximum x-coordinate of the virtual splitter"
Top             =   9255
Width           =   915
End
Begin VB.ComboBox cboSplLiveUpdate
Height          =   315
ItemData        =   "frmDemoFeatures.frx":00B0
Left            =   1605
List            =   "frmDemoFeatures.frx":00BA
Style           =   2                             'Dropdown List
TabIndex        =   108
ToolTipText     =   "Returns/sets a value that determines whether the controls should be resized as the splitter is moved"
Top             =   8745
Width           =   915
End
Begin VB.TextBox txtSplLeft
Enabled         =   0                             'False
Height          =   285
Left            =   1605
TabIndex        =   107
ToolTipText     =   $"frmDemoFeatures.frx":00CB
Top             =   8280
Width           =   915
End
Begin VB.ListBox lstSplIdsSplTop
Enabled         =   0                             'False
Height          =   450
Left            =   1605
TabIndex        =   106
ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's up-movement"
Top             =   7661
Width           =   900
End
Begin VB.ListBox lstSplIdsSplRight
Enabled         =   0                             'False
Height          =   450
Left            =   1605
TabIndex        =   105
ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's right-movement"
Top             =   7042
Width           =   900
End
Begin VB.ListBox lstSplIdsSplLeft
Enabled         =   0                             'False
Height          =   450
Left            =   1605
TabIndex        =   104
ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's left-movement"
Top             =   6423
Width           =   900
End
Begin VB.ListBox lstSplIdsSplBottom
Enabled         =   0                             'False
Height          =   450
Left            =   1605
TabIndex        =   103
ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's down-movement"
Top             =   5804
Width           =   900
End
Begin VB.ListBox lstSplIdsCtlTop
Enabled         =   0                             'False
Height          =   450
Left            =   1605
TabIndex        =   102
ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's up-movement"
Top             =   5185
Width           =   900
End
Begin VB.ListBox lstSplIdsCtlRight
Enabled         =   0                             'False
Height          =   450
Left            =   1605
TabIndex        =   101
ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's right-movement"
Top             =   4566
Width           =   900
End
Begin VB.ListBox lstSplIdsCtlLeft
Enabled         =   0                             'False
Height          =   450
Left            =   1605
TabIndex        =   100
ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's left-movement"
Top             =   3947
Width           =   900
End
Begin VB.ListBox lstSplIdsCtlBottom
Enabled         =   0                             'False
Height          =   450
Left            =   1605
TabIndex        =   99
ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's down-movement"
Top             =   3328
Width           =   900
End
Begin VB.TextBox txtSplId
Enabled         =   0                             'False
Height          =   285
Left            =   1605
TabIndex        =   98
ToolTipText     =   "Returns a value that uniquely identifies the virtual splitter"
Top             =   2874
Width           =   915
End
Begin VB.TextBox txtSplHeight
Enabled         =   0                             'False
Height          =   285
Left            =   1605
TabIndex        =   97
ToolTipText     =   "Returns the height of the virtual splitter"
Top             =   2420
Width           =   915
End
Begin VB.ComboBox cboSplEnable
Height          =   315
ItemData        =   "frmDemoFeatures.frx":0152
Left            =   1605
List            =   "frmDemoFeatures.frx":015C
Style           =   2                             'Dropdown List
TabIndex        =   96
ToolTipText     =   "Returns/sets a value that determines whether the splitter is movable"
Top             =   1936
Width           =   915
End
Begin VB.ComboBox cboSplClipCursor
Height          =   315
ItemData        =   "frmDemoFeatures.frx":016D
Left            =   1605
List            =   "frmDemoFeatures.frx":0177
Style           =   2                             'Dropdown List
TabIndex        =   95
ToolTipText     =   $"frmDemoFeatures.frx":0188
Top             =   1452
Width           =   915
End
Begin VB.TextBox txtSplBottom
Enabled         =   0                             'False
Height          =   285
Left            =   1605
TabIndex        =   94
ToolTipText     =   $"frmDemoFeatures.frx":0237
Top             =   998
Width           =   915
End
Begin VB.Label lblSplClick
Caption         =   "(click the splitter)"
BeginProperty Font
Name            =   "MS Sans Serif"
Size            =   8.25
Charset         =   0
Weight          =   400
Underline       =   -1                            'True
Italic          =   0                             'False
Strikethrough   =   0                             'False
EndProperty
Height          =   240
Left            =   315
TabIndex        =   148
Top             =   2970
Width           =   1245
End
Begin VB.Label Label76
AutoSize        =   -1                            'True
Caption         =   "Orientation:"
Height          =   195
Left            =   105
TabIndex        =   147
ToolTipText     =   "Returns the virtual splitter movement direction"
Top             =   11475
Width           =   810
End
Begin VB.Label Label75
AutoSize        =   -1                            'True
Caption         =   "Yc:"
Height          =   195
Left            =   105
TabIndex        =   146
ToolTipText     =   "Returns the y-coordinate of the virtual splitter center"
Top             =   14160
Width           =   240
End
Begin VB.Label Label74
AutoSize        =   -1                            'True
Caption         =   "Xc:"
Height          =   195
Left            =   105
TabIndex        =   145
ToolTipText     =   "Returns the x-coordinate of the virtual splitter center"
Top             =   13620
Width           =   240
End
Begin VB.Label Label73
AutoSize        =   -1                            'True
Caption         =   "Width:"
Height          =   195
Left            =   105
TabIndex        =   144
ToolTipText     =   "Returns the width of the virtual splitter"
Top             =   13080
Width           =   465
End
Begin VB.Label Label72
AutoSize        =   -1                            'True
Caption         =   "Top:"
Height          =   195
Left            =   105
TabIndex        =   143
ToolTipText     =   "Returns the distance between the internal top edge of the virtual splitter and the top edge of the related Control Manager object"
Top             =   12540
Width           =   330
End
Begin VB.Label Label71
AutoSize        =   -1                            'True
Caption         =   "Right:"
Height          =   195
Left            =   105
TabIndex        =   142
ToolTipText     =   $"frmDemoFeatures.frx":02BF
Top             =   12015
Width           =   420
End
Begin VB.Label Label70
AutoSize        =   -1                            'True
Caption         =   "MinYc:"
Height          =   195
Left            =   105
TabIndex        =   141
ToolTipText     =   "Returns the minimum y-coordinate of the virtual splitter"
Top             =   10935
Width           =   495
End
Begin VB.Label Label69
AutoSize        =   -1                            'True
Caption         =   "MinXc:"
Height          =   195
Left            =   105
TabIndex        =   140
ToolTipText     =   "Returns the minimum x-coordinate of the virtual splitter"
Top             =   10395
Width           =   495
End
Begin VB.Label Label68
AutoSize        =   -1                            'True
Caption         =   "MaxYc:"
Height          =   195
Left            =   105
TabIndex        =   139
ToolTipText     =   "Returns the maximum y-coordinate of the virtual splitter"
Top             =   9870
Width           =   540
End
Begin VB.Label Label67
AutoSize        =   -1                            'True
Caption         =   "MaxXc:"
Height          =   195
Left            =   105
TabIndex        =   138
ToolTipText     =   "Returns the maximum x-coordinate of the virtual splitter"
Top             =   9330
Width           =   540
End
Begin VB.Label Label66
AutoSize        =   -1                            'True
Caption         =   "LiveUpdate:"
Height          =   195
Left            =   105
TabIndex        =   137
ToolTipText     =   "Returns/sets a value that determines whether the controls should be resized as the splitter is moved"
Top             =   8835
Width           =   870
End
Begin VB.Label Label65
AutoSize        =   -1                            'True
Caption         =   "Left:"
Height          =   195
Left            =   105
TabIndex        =   136
ToolTipText     =   $"frmDemoFeatures.frx":0347
Top             =   8370
Width           =   315
End
Begin VB.Label Label64
AutoSize        =   -1                            'True
Caption         =   "IdsCtlTop:"
Height          =   195
Left            =   105
TabIndex        =   135
ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's up-movement"
Top             =   7650
Width           =   720
End
Begin VB.Label Label63
AutoSize        =   -1                            'True
Caption         =   "IdsSplRight:"
Height          =   195
Left            =   105
TabIndex        =   134
ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's right-movement"
Top             =   7035
Width           =   855
End
Begin VB.Label Label62
AutoSize        =   -1                            'True
Caption         =   "IdsSplLeft:"
Height          =   195
Left            =   105
TabIndex        =   133
ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's left-movement"
Top             =   6435
Width           =   750
End
Begin VB.Label Label61
AutoSize        =   -1                            'True
Caption         =   "IdsSplBottom:"
Height          =   195
Left            =   105
TabIndex        =   132
ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's down-movement"
Top             =   5820
Width           =   975
End
Begin VB.Label Label60
AutoSize        =   -1                            'True
Caption         =   "IdsCtlTop:"
Height          =   195
Left            =   120
TabIndex        =   131
ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's up-movement"
Top             =   5220
Width           =   720
End
Begin VB.Label Label59
AutoSize        =   -1                            'True
Caption         =   "IdsCtlRight:"
Height          =   195
Left            =   105
TabIndex        =   130
ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's right-movement"
Top             =   4605
Width           =   810
End
Begin VB.Label Label58
AutoSize        =   -1                            'True
Caption         =   "IdsCtlLeft:"
Height          =   195
Left            =   105
TabIndex        =   129
ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's left-movement"
Top             =   4005
Width           =   705
End
Begin VB.Label Label57
AutoSize        =   -1                            'True
Caption         =   "IdsCtlBottom:"
Height          =   195
Left            =   105
TabIndex        =   128
ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's down-movement"
Top             =   3390
Width           =   930
End
Begin VB.Label Label56
AutoSize        =   -1                            'True
Caption         =   "Id:"
Height          =   195
Left            =   105
TabIndex        =   127
ToolTipText     =   "Returns a value that uniquely identifies the virtual splitter"
Top             =   2985
Width           =   180
End
Begin VB.Label Label54
AutoSize        =   -1                            'True
Caption         =   "Height:"
Height          =   195
Left            =   105
TabIndex        =   126
ToolTipText     =   "Returns the height of the virtual splitter"
Top             =   2505
Width           =   510
End
Begin VB.Label Label21
AutoSize        =   -1                            'True
Caption         =   "Enable:"
Height          =   195
Left            =   105
TabIndex        =   125
ToolTipText     =   "Returns/sets a value that determines whether the splitter is movable"
Top             =   2040
Width           =   540
End
Begin VB.Label Label55
AutoSize        =   -1                            'True
Caption         =   "BackColor:"
Height          =   195
Left            =   105
TabIndex        =   124
ToolTipText     =   "Returns/sets the background color used to display the splitter"
Top             =   615
Width           =   780
End
Begin VB.Label lblSplBackColor
BackColor       =   &H00404040&
BorderStyle     =   1                             'Fixed Single
Height          =   255
Left            =   1605
TabIndex        =   123
ToolTipText     =   "Returns/sets the background color used to display the splitter"
Top             =   574
Width           =   915
End
Begin VB.Label Label53
AutoSize        =   -1                            'True
Caption         =   "Active Color:"
Height          =   195
Left            =   105
TabIndex        =   122
ToolTipText     =   "Returns/sets the background color used to display the splitter when the user moves it in none live update mode"
Top             =   150
Width           =   900
End
Begin VB.Label lblSplActiveColor
BackColor       =   &H00404040&
BorderStyle     =   1                             'Fixed Single
Height          =   255
Left            =   1605
TabIndex        =   121
ToolTipText     =   "Returns/sets the background color used to display the splitter when the user moves it in none live update mode"
Top             =   150
Width           =   915
End
Begin VB.Label Label20
AutoSize        =   -1                            'True
Caption         =   "ClipCursor:"
Height          =   195
Left            =   105
TabIndex        =   120
ToolTipText     =   $"frmDemoFeatures.frx":03CE
Top             =   1560
Width           =   750
End
Begin VB.Label Label19
AutoSize        =   -1                            'True
Caption         =   "Bottom:"
Height          =   195
Left            =   105
TabIndex        =   119
ToolTipText     =   $"frmDemoFeatures.frx":047D
Top             =   1095
Width           =   540
End
End
End
End
Begin VB.Frame fraFeatures
Height          =   7680
Index           =   3
Left            =   360
TabIndex        =   90
Top             =   420
Width           =   3030
Begin VB.VScrollBar vsbCtlProperties
Height          =   7530
Left            =   2820
TabIndex        =   151
Top             =   0
Width           =   210
End
Begin VB.Frame fraConCtlProperties
BackColor       =   &H8000000B&
BorderStyle     =   0                             'None
Caption         =   "Frame1"
Height          =   6435
Left            =   210
TabIndex        =   150
Top             =   240
Width           =   2430
Begin VB.Frame fraCtlProperties
BorderStyle     =   0                             'None
Caption         =   "Frame1"
Height          =   7425
Left            =   510
TabIndex        =   152
Top             =   480
Width           =   2745
Begin VB.TextBox txtCtlYc
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   172
ToolTipText     =   "Returns the y-coordinate of the virtual control center"
Top             =   8355
Width           =   915
End
Begin VB.TextBox txtCtlXc
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   171
ToolTipText     =   "Returns the x-coordinate of the virtual control center"
Top             =   7920
Width           =   915
End
Begin VB.TextBox txtCtlWidth
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   170
ToolTipText     =   "Returns the width of the virtual control"
Top             =   7485
Width           =   915
End
Begin VB.TextBox txtCtlTop
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   169
ToolTipText     =   "Returns the distance between the internal top edge of the virtual control and the top edge of the related Control Manager object"
Top             =   7065
Width           =   915
End
Begin VB.TextBox txtCtlRight
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   168
ToolTipText     =   $"frmDemoFeatures.frx":0505
Top             =   5280
Width           =   915
End
Begin VB.TextBox txtCtlName
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   167
ToolTipText     =   "Returns the name of the real control that the virtual control represents"
Top             =   4845
Width           =   915
End
Begin VB.TextBox txtCtlMinWidth
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   166
ToolTipText     =   "Returns the minimum width of the virtual control"
Top             =   4425
Width           =   915
End
Begin VB.TextBox txtCtlMinHeight
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   165
ToolTipText     =   "Returns the minimum height of the virtual control"
Top             =   3990
Width           =   915
End
Begin VB.TextBox txtCtlLeft
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   164
ToolTipText     =   "Returns the distance between the internal left edge of the virtual control and the left edge of the related Control Manager object"
Top             =   3555
Width           =   915
End
Begin VB.TextBox txtCtlIdSplTop
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   163
ToolTipText     =   $"frmDemoFeatures.frx":058C
Top             =   3135
Width           =   915
End
Begin VB.TextBox txtCtlIdSplRight
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   162
ToolTipText     =   $"frmDemoFeatures.frx":061C
Top             =   2700
Width           =   915
End
Begin VB.TextBox txtCtlIdSplLeft
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   161
ToolTipText     =   $"frmDemoFeatures.frx":06AE
Top             =   2265
Width           =   915
End
Begin VB.TextBox txtCtlIdSplBottom
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   160
ToolTipText     =   $"frmDemoFeatures.frx":073F
Top             =   1845
Width           =   915
End
Begin VB.TextBox txtCtlId
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   159
ToolTipText     =   "Returns a value that uniquely identifies the virtual control"
Top             =   1410
Width           =   915
End
Begin VB.TextBox txtCtlHeight
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   158
ToolTipText     =   "Returns the height of the virtual control"
Top             =   975
Width           =   915
End
Begin VB.ComboBox cboCtlClosed
Enabled         =   0                             'False
Height          =   315
ItemData        =   "frmDemoFeatures.frx":07D2
Left            =   1680
List            =   "frmDemoFeatures.frx":07DC
Style           =   2                             'Dropdown List
TabIndex        =   157
ToolTipText     =   "Returns a value that determines whether the virtual control is closed"
Top             =   525
Width           =   915
End
Begin VB.TextBox txtCtlBottom
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   156
ToolTipText     =   $"frmDemoFeatures.frx":07ED
Top             =   90
Width           =   915
End
Begin VB.TextBox txtCtlTitleBarHeight
Enabled         =   0                             'False
Height          =   285
Left            =   1680
TabIndex        =   155
ToolTipText     =   "Returns the height of the virtual control title bars"
Top             =   6165
Width           =   915
End
Begin VB.ComboBox cboCtlTitleBarCloseVisible
Height          =   315
ItemData        =   "frmDemoFeatures.frx":0874
Left            =   1680
List            =   "frmDemoFeatures.frx":087E
Style           =   2                             'Dropdown List
TabIndex        =   154
ToolTipText     =   "Returns/sets a value that determines whether a close button in the virtual control title bar is visible"
Top             =   5715
Width           =   915
End
Begin VB.ComboBox cboCtlTitleBarVisible
Height          =   315
ItemData        =   "frmDemoFeatures.frx":088F
Left            =   1680
List            =   "frmDemoFeatures.frx":0899
Style           =   2                             'Dropdown List
TabIndex        =   153
ToolTipText     =   "Returns/sets a value that determines whether the virtual control title bar is visible"
Top             =   6600
Width           =   915
End
Begin VB.Label Label52
AutoSize        =   -1                            'True
Caption         =   "Yc:"
Height          =   195
Left            =   105
TabIndex        =   193
ToolTipText     =   "Returns the y-coordinate of the virtual control center"
Top             =   8430
Width           =   240
End
Begin VB.Label Label51
AutoSize        =   -1                            'True
Caption         =   "Xc:"
Height          =   195
Left            =   105
TabIndex        =   192
ToolTipText     =   "Returns the x-coordinate of the virtual control center"
Top             =   7980
Width           =   240
End
Begin VB.Label Label50
AutoSize        =   -1                            'True
Caption         =   "Width:"
Height          =   195
Left            =   105
TabIndex        =   191
ToolTipText     =   "Returns the width of the virtual control"
Top             =   7545
Width           =   465
End
Begin VB.Label Label49
AutoSize        =   -1                            'True
Caption         =   "Top:"
Height          =   195
Left            =   105
TabIndex        =   190
ToolTipText     =   "Returns the distance between the internal top edge of the virtual control and the top edge of the related Control Manager object"
Top             =   7110
Width           =   330
End
Begin VB.Label Label48
AutoSize        =   -1                            'True
Caption         =   "Right:"
Height          =   195
Left            =   105
TabIndex        =   189
ToolTipText     =   $"frmDemoFeatures.frx":08AA
Top             =   5385
Width           =   420
End
Begin VB.Label Label47
AutoSize        =   -1                            'True
Caption         =   "Name:"
Height          =   195
Left            =   105
TabIndex        =   188
ToolTipText     =   "Returns the name of the real control that the virtual control represents"
Top             =   4935
Width           =   465
End
Begin VB.Label Label46
AutoSize        =   -1                            'True
Caption         =   "MinWidth:"
Height          =   195
Left            =   105
TabIndex        =   187
ToolTipText     =   "Returns the minimum width of the virtual control"
Top             =   4500
Width           =   720
End
Begin VB.Label Label45
AutoSize        =   -1                            'True
Caption         =   "MinHeight:"
Height          =   195
Left            =   105
TabIndex        =   186
ToolTipText     =   "Returns the minimum height of the virtual control"
Top             =   4065
Width           =   765
End
Begin VB.Label Label44
AutoSize        =   -1                            'True
Caption         =   "Left:"
Height          =   195
Left            =   105
TabIndex        =   185
ToolTipText     =   "Returns the distance between the internal left edge of the virtual control and the left edge of the related Control Manager object"
Top             =   3630
Width           =   315
End
Begin VB.Label Label43
AutoSize        =   -1                            'True
Caption         =   "IdSplTop:"
Height          =   195
Left            =   105
TabIndex        =   184
ToolTipText     =   $"frmDemoFeatures.frx":0931
Top             =   3210
Width           =   690
End
Begin VB.Label Label42
AutoSize        =   -1                            'True
Caption         =   "IdSplRight:"
Height          =   195
Left            =   105
TabIndex        =   183
ToolTipText     =   $"frmDemoFeatures.frx":09C1
Top             =   2775
Width           =   780
End
Begin VB.Label Label41
AutoSize        =   -1                            'True
Caption         =   "IdSplLeft:"
Height          =   195
Left            =   105
TabIndex        =   182
ToolTipText     =   $"frmDemoFeatures.frx":0A53
Top             =   2340
Width           =   675
End
Begin VB.Label Label40
AutoSize        =   -1                            'True
Caption         =   "IdSplBottom:"
Height          =   195
Left            =   105
TabIndex        =   181
ToolTipText     =   $"frmDemoFeatures.frx":0AE4
Top             =   1905
Width           =   900
End
Begin VB.Label Label39
AutoSize        =   -1                            'True
Caption         =   "Id:"
Height          =   195
Left            =   105
TabIndex        =   180
ToolTipText     =   "Returns a value that uniquely identifies the virtual control"
Top             =   1470
Width           =   180
End
Begin VB.Label Label38
AutoSize        =   -1                            'True
Caption         =   "Height:"
Height          =   195
Left            =   105
TabIndex        =   179
ToolTipText     =   "Returns the height of the virtual control"
Top             =   1035
Width           =   510
End
Begin VB.Label Label37
AutoSize        =   -1                            'True
Caption         =   "Closed:"
Height          =   195
Left            =   105
TabIndex        =   178
ToolTipText     =   "Returns a value that determines whether the virtual control is closed"
Top             =   600
Width           =   525
End
Begin VB.Label Label16
AutoSize        =   -1                            'True
Caption         =   "Bottom:"
Height          =   195
Left            =   105
TabIndex        =   177
ToolTipText     =   $"frmDemoFeatures.frx":0B77
Top             =   165
Width           =   540
End
Begin VB.Label Label36
AutoSize        =   -1                            'True
Caption         =   "TitleBar_CloseVisible:"
Height          =   195
Left            =   105
TabIndex        =   176
ToolTipText     =   "Returns/sets a value that determines whether a close button in the virtual control title bar is visible"
Top             =   5820
Width           =   1515
End
Begin VB.Label Label35
AutoSize        =   -1                            'True
Caption         =   "TitleBar_Height:"
Height          =   195
Left            =   105
TabIndex        =   175
ToolTipText     =   "Returns the height of the virtual control title bars"
Top             =   6240
Width           =   1140
End
Begin VB.Label Label34
AutoSize        =   -1                            'True
Caption         =   "TitleBar_Visible:"
Height          =   195
Left            =   105
TabIndex        =   174
ToolTipText     =   "Returns/sets a value that determines whether the virtual control title bar is visible"
Top             =   6675
Width           =   1125
End
Begin VB.Label lblCtlClick
Caption         =   "(click the title bar)"
BeginProperty Font
Name            =   "MS Sans Serif"
Size            =   8.25
Charset         =   0
Weight          =   400
Underline       =   -1                            'True
Italic          =   0                             'False
Strikethrough   =   0                             'False
EndProperty
Height          =   225
Left            =   330
TabIndex        =   173
Top             =   1470
Width           =   1275
End
End
End
End
Begin VB.Frame fraFeatures
Height          =   7695
Index           =   0
Left            =   270
TabIndex        =   61
Top             =   600
Width           =   3030
Begin VB.ComboBox cboCMTitleBarVisible
Height          =   315
ItemData        =   "frmDemoFeatures.frx":0BFE
Left            =   1725
List            =   "frmDemoFeatures.frx":0C08
Style           =   2                             'Dropdown List
TabIndex        =   73
ToolTipText     =   "Returns/sets a value that determines whether all control title bars are visible"
Top             =   5625
Width           =   915
End
Begin VB.ComboBox cboCMTitleBarCloseVisible
Height          =   315
ItemData        =   "frmDemoFeatures.frx":0C19
Left            =   1725
List            =   "frmDemoFeatures.frx":0C23
Style           =   2                             'Dropdown List
TabIndex        =   72
ToolTipText     =   "Returns/sets a value that determines whether a close button in all control title bars is visible"
Top             =   4770
Width           =   915
End
Begin VB.TextBox txtCMTitleBarHeight
Enabled         =   0                             'False
Height          =   285
Left            =   1725
TabIndex        =   71
ToolTipText     =   "Returns/sets the height of all control title bars"
Top             =   5220
Width           =   915
End
Begin VB.TextBox txtCMSize
Height          =   285
Left            =   1725
TabIndex        =   70
ToolTipText     =   "Returns/sets the size of all splitters"
Top             =   4365
Width           =   915
End
Begin VB.TextBox txtCMMarginTop
Height          =   285
Left            =   1725
TabIndex        =   69
ToolTipText     =   "Returns/sets the top margin of the ActiveX Control from its container"
Top             =   3960
Width           =   915
End
Begin VB.TextBox txtCMMarginRight
Height          =   285
Left            =   1725
TabIndex        =   68
ToolTipText     =   "Returns/sets the right margin of the ActiveX Control from its container"
Top             =   3555
Width           =   915
End
Begin VB.TextBox txtCMMarginLeft
Height          =   285
Left            =   1740
TabIndex        =   67
ToolTipText     =   "Returns/sets the left margin of the ActiveX Control from its container"
Top             =   3150
Width           =   915
End
Begin VB.TextBox txtCMMarginBottom
Height          =   285
Left            =   1725
TabIndex        =   66
ToolTipText     =   "Returns/sets the bottom margin of the ActiveX Control from its container"
Top             =   2745
Width           =   915
End
Begin VB.ComboBox cboCMLiveUpdate
Height          =   315
ItemData        =   "frmDemoFeatures.frx":0C34
Left            =   1725
List            =   "frmDemoFeatures.frx":0C3E
Style           =   2                             'Dropdown List
TabIndex        =   65
ToolTipText     =   "Returns/sets a value that determines whether the controls should be resized as a splitter is moved"
Top             =   2310
Width           =   915
End
Begin VB.ComboBox cboCMEnable
Height          =   315
ItemData        =   "frmDemoFeatures.frx":0C4F
Left            =   1725
List            =   "frmDemoFeatures.frx":0C59
Style           =   2                             'Dropdown List
TabIndex        =   64
ToolTipText     =   "Returns/sets a value that determines whether all splitters are movable"
Top             =   1425
Width           =   915
End
Begin VB.ComboBox cboCMFillContainer
Height          =   315
ItemData        =   "frmDemoFeatures.frx":0C6A
Left            =   1725
List            =   "frmDemoFeatures.frx":0C74
Style           =   2                             'Dropdown List
TabIndex        =   63
ToolTipText     =   $"frmDemoFeatures.frx":0C85
Top             =   1875
Width           =   915
End
Begin VB.ComboBox cboCMClipCursor
Height          =   315
ItemData        =   "frmDemoFeatures.frx":0D4F
Left            =   1725
List            =   "frmDemoFeatures.frx":0D59
Style           =   2                             'Dropdown List
TabIndex        =   62
ToolTipText     =   $"frmDemoFeatures.frx":0D6A
Top             =   1005
Width           =   915
End
Begin VB.Label Label24
AutoSize        =   -1                            'True
Caption         =   "TitleBar_Visible:"
Height          =   195
Left            =   150
TabIndex        =   89
ToolTipText     =   "Returns/sets a value that determines whether all control title bars are visible"
Top             =   5715
Width           =   1125
End
Begin VB.Label Label23
AutoSize        =   -1                            'True
Caption         =   "TitleBar_Height:"
Height          =   195
Left            =   150
TabIndex        =   88
ToolTipText     =   "Returns/sets the height of all control title bars"
Top             =   5295
Width           =   1140
End
Begin VB.Label Label22
AutoSize        =   -1                            'True
Caption         =   "TitleBar_CloseVisible:"
Height          =   195
Left            =   135
TabIndex        =   87
ToolTipText     =   "Returns/sets a value that determines whether a close button in all control title bars is visible"
Top             =   4875
Width           =   1515
End
Begin VB.Label Label13
AutoSize        =   -1                            'True
Caption         =   "Size:"
Height          =   195
Left            =   150
TabIndex        =   86
ToolTipText     =   "Returns/sets the size of all splitters"
Top             =   4455
Width           =   345
End
Begin VB.Label Label12
AutoSize        =   -1                            'True
Caption         =   "MarginTop:"
Height          =   195
Left            =   150
TabIndex        =   85
ToolTipText     =   "Returns/sets the top margin of the ActiveX Control from its container"
Top             =   4050
Width           =   810
End
Begin VB.Label Label11
AutoSize        =   -1                            'True
Caption         =   "MarginRight:"
Height          =   195
Left            =   150
TabIndex        =   84
ToolTipText     =   "Returns/sets the right margin of the ActiveX Control from its container"
Top             =   3630
Width           =   900
End
Begin VB.Label Label10
AutoSize        =   -1                            'True
Caption         =   "MarginLeft:"
Height          =   195
Left            =   150
TabIndex        =   83
ToolTipText     =   "Returns/sets the left margin of the ActiveX Control from its container"
Top             =   3210
Width           =   795
End
Begin VB.Label Label9
AutoSize        =   -1                            'True
Caption         =   "MarginBottom:"
Height          =   195
Left            =   150
TabIndex        =   82
ToolTipText     =   "Returns/sets the bottom margin of the ActiveX Control from its container"
Top             =   2790
Width           =   1020
End
Begin VB.Label Label8
AutoSize        =   -1                            'True
Caption         =   "Live Update:"
Height          =   195
Left            =   150
TabIndex        =   81
ToolTipText     =   "Returns/sets a value that determines whether the controls should be resized as a splitter is moved"
Top             =   2385
Width           =   915
End
Begin VB.Label Label7
AutoSize        =   -1                            'True
Caption         =   "Enable:"
Height          =   195
Left            =   150
TabIndex        =   80
ToolTipText     =   "Returns/sets a value that determines whether all splitters are movable"
Top             =   1545
Width           =   540
End
Begin VB.Label Label6
AutoSize        =   -1                            'True
Caption         =   "Fill Container:"
Height          =   195
Left            =   150
TabIndex        =   79
ToolTipText     =   $"frmDemoFeatures.frx":0E17
Top             =   1965
Width           =   945
End
Begin VB.Label Label5
AutoSize        =   -1                            'True
Caption         =   "Clip Cursor:"
Height          =   195
Left            =   150
TabIndex        =   78
ToolTipText     =   $"frmDemoFeatures.frx":0EE1
Top             =   1125
Width           =   795
End
Begin VB.Label lblCMBackColor
BorderStyle     =   1                             'Fixed Single
Height          =   255
Left            =   1725
TabIndex        =   77
ToolTipText     =   "Returns/sets the background color used to display all splitters"
Top             =   615
Width           =   915
End
Begin VB.Label lblCMActiveColor
BackColor       =   &H00404040&
BorderStyle     =   1                             'Fixed Single
Height          =   255
Left            =   1740
TabIndex        =   76
ToolTipText     =   "Returns/sets the background color used to display a splitter when the user moves it in none live update mode"
Top             =   240
Width           =   915
End
Begin VB.Label Label2
AutoSize        =   -1                            'True
Caption         =   "Active Color:"
Height          =   195
Left            =   150
TabIndex        =   75
ToolTipText     =   "Returns/sets the background color used to display a splitter when the user moves it in none live update mode"
Top             =   300
Width           =   900
End
Begin VB.Label Label1
AutoSize        =   -1                            'True
Caption         =   "Back Color:"
Height          =   195
Left            =   150
TabIndex        =   74
ToolTipText     =   "Returns/sets the background color used to display all splitters"
Top             =   720
Width           =   825
End
End
Begin VB.Frame fraFeatures
Height          =   7680
Index           =   1
Left            =   300
TabIndex        =   34
Top             =   420
Width           =   3030
Begin VB.ComboBox cboOpenControlMaintainSize
Height          =   315
ItemData        =   "frmDemoFeatures.frx":0F8E
Left            =   1275
List            =   "frmDemoFeatures.frx":0F98
Style           =   2                             'Dropdown List
TabIndex        =   41
Top             =   4755
Width           =   1485
End
Begin VB.ComboBox cboMoveControlMoveTo
Height          =   315
ItemData        =   "frmDemoFeatures.frx":0FA9
Left            =   1005
List            =   "frmDemoFeatures.frx":0FC8
Style           =   2                             'Dropdown List
TabIndex        =   40
Top             =   1710
Width           =   1710
End
Begin VB.CommandButton cmdMoveControl
Caption         =   "Call"
Height          =   285
Left            =   2070
TabIndex        =   39
Top             =   1005
Width           =   660
End
Begin VB.CommandButton cmdOpenControl
Caption         =   "Call"
Height          =   285
Left            =   2070
TabIndex        =   38
Top             =   4080
Width           =   660
End
Begin VB.CommandButton cmdCloseControl
Caption         =   "Call"
Height          =   285
Left            =   2070
TabIndex        =   37
Top             =   240
Width           =   660
End
Begin VB.CommandButton cmdMoveSplitter
Caption         =   "Call"
Height          =   285
Left            =   2070
TabIndex        =   36
Top             =   2910
Width           =   660
End
Begin VB.TextBox txtMoveSplitterMoveTo
Height          =   300
Left            =   990
TabIndex        =   35
Top             =   3615
Width           =   1725
End
Begin VB.Label Label33
Caption         =   "IdSplitterDesination:"
Height          =   255
Left            =   225
TabIndex        =   60
Tag             =   $"frmDemoFeatures.frx":1048
Top             =   2565
Width           =   1410
End
Begin VB.Label lblIdSplitter
Caption         =   "(click the spliter)"
Height          =   255
Index           =   0
Left            =   1695
TabIndex        =   59
Top             =   2565
Width           =   1200
End
Begin VB.Label lblIdControl2
Caption         =   "(double-click the title bar)"
Height          =   435
Left            =   1695
TabIndex        =   58
Top             =   2175
Width           =   1200
End
Begin VB.Label Label17
Caption         =   "IdControlDesination:"
Height          =   255
Left            =   210
TabIndex        =   57
ToolTipText     =   $"frmDemoFeatures.frx":10EB
Top             =   2175
Width           =   1410
End
Begin VB.Label Label29
Caption         =   "MaintainSize:"
Height          =   255
Left            =   195
TabIndex        =   56
ToolTipText     =   $"frmDemoFeatures.frx":1199
Top             =   4860
Width           =   1050
End
Begin VB.Label Label32
Caption         =   "IdControlSource:"
Height          =   255
Left            =   195
TabIndex        =   55
ToolTipText     =   "Required. A value that uniquely identifies the source control the developer want to move"
Top             =   1410
Width           =   1260
End
Begin VB.Label Label31
Caption         =   "MoveTo:"
Height          =   255
Left            =   210
TabIndex        =   54
ToolTipText     =   "Required. A value indicating the area type where the source control will be moved to"
Top             =   1800
Width           =   810
End
Begin VB.Label lblIdControl
Caption         =   "(click the title bar)"
Height          =   255
Index           =   1
Left            =   1455
TabIndex        =   53
Top             =   1410
Width           =   1335
End
Begin VB.Label Label26
Caption         =   "MoveControl"
BeginProperty Font
Name            =   "MS Sans Serif"
Size            =   8.25
Charset         =   0
Weight          =   700
Underline       =   0                             'False
Italic          =   0                             'False
Strikethrough   =   0                             'False
EndProperty
Height          =   270
Left            =   195
TabIndex        =   52
ToolTipText     =   "Moves a control to certain area"
Top             =   1050
Width           =   1590
End
Begin VB.Label Label30
Caption         =   "IdControl:"
Height          =   255
Left            =   195
TabIndex        =   51
ToolTipText     =   "Required. A value that uniquely identifies the control the developer want to open"
Top             =   4485
Width           =   705
End
Begin VB.Label lblIdControl
Caption         =   "(click the title bar)"
Height          =   255
Index           =   2
Left            =   975
TabIndex        =   50
Top             =   4485
Width           =   1800
End
Begin VB.Label Label28
Caption         =   "OpenControl"
BeginProperty Font
Name            =   "MS Sans Serif"
Size            =   8.25
Charset         =   0
Weight          =   700
Underline       =   0                             'False
Italic          =   0                             'False
Strikethrough   =   0                             'False
EndProperty
Height          =   270
Left            =   195
TabIndex        =   49
ToolTipText     =   "Opens (shows) a control and docks it to the ActiveX Control"
Top             =   4125
Width           =   1590
End
Begin VB.Label Label27
Caption         =   "IdControl:"
Height          =   255
Left            =   195
TabIndex        =   48
ToolTipText     =   "A value that uniquely identifies the control the developer want to close"
Top             =   645
Width           =   705
End
Begin VB.Label lblIdControl
Caption         =   "(click the title bar)"
Height          =   255
Index           =   0
Left            =   1005
TabIndex        =   47
Top             =   645
Width           =   1800
End
Begin VB.Label Label25
Caption         =   "CloseControl"
BeginProperty Font
Name            =   "MS Sans Serif"
Size            =   8.25
Charset         =   0
Weight          =   700
Underline       =   0                             'False
Italic          =   0                             'False
Strikethrough   =   0                             'False
EndProperty
Height          =   270
Left            =   195
TabIndex        =   46
ToolTipText     =   "Closes (hides) a control"
Top             =   285
Width           =   1590
End
Begin VB.Label Label15
Caption         =   "MoveSplitter"
BeginProperty Font
Name            =   "MS Sans Serif"
Size            =   8.25
Charset         =   0
Weight          =   700
Underline       =   0                             'False
Italic          =   0                             'False
Strikethrough   =   0                             'False
EndProperty
Height          =   270
Left            =   195
TabIndex        =   45
ToolTipText     =   "Moves a splitter to the specified x- or y- (depending on the splitter's Orientation property) coordinate"
Top             =   2955
Width           =   1590
End
Begin VB.Label lblIdSplitter
Caption         =   "(click the spliter)"
Height          =   255
Index           =   1
Left            =   975
TabIndex        =   44
Top             =   3315
Width           =   1800
End
Begin VB.Label Label4
Caption         =   "MoveTo:"
Height          =   255
Left            =   195
TabIndex        =   43
ToolTipText     =   $"frmDemoFeatures.frx":123D
Top             =   3705
Width           =   810
End
Begin VB.Label Label3
Caption         =   "IdSplitter:"
Height          =   255
Left            =   195
TabIndex        =   42
ToolTipText     =   "Required. A value that uniquely identifies the splitter the developer want to move"
Top             =   3315
Width           =   705
End
End
Begin VB.Frame fraFeatures
Height          =   7665
Index           =   2
Left            =   330
TabIndex        =   15
Top             =   420
Width           =   3030
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "TitleBarMouseUp"
Height          =   255
Index           =   17
Left            =   225
TabIndex        =   33
ToolTipText     =   "Occurs when the user releases a mouse button over a control title bar without previously moving the control"
Top             =   6645
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "TitleBarMouseMove"
Height          =   255
Index           =   16
Left            =   225
TabIndex        =   32
ToolTipText     =   "Occurs when the user moves a mouse over a control title bar without moving the control"
Top             =   6285
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "TitleBarMouseDown"
Height          =   225
Index           =   15
Left            =   225
TabIndex        =   31
ToolTipText     =   "Occurs when the user presses a mouse button over a control title bar"
Top             =   5925
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "TitleBarDblClick"
Height          =   255
Index           =   14
Left            =   225
TabIndex        =   30
ToolTipText     =   "Occurs when the user presses and then realeses a mouse button and then presses and releases it again over a control title bar"
Top             =   5550
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "TitleBarClick"
Height          =   255
Index           =   13
Left            =   225
TabIndex        =   29
ToolTipText     =   "Occurs when the user presses and then realeses a mouse button over a control title bar"
Top             =   5175
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "SplitterMoveEnd"
Height          =   255
Index           =   12
Left            =   225
TabIndex        =   28
ToolTipText     =   "Occurs when the user presses and then realeses a mouse button over a control title bar"
Top             =   4800
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "SplitterMoveBegin"
Height          =   255
Index           =   11
Left            =   225
TabIndex        =   27
ToolTipText     =   "Occurs when the user is about to move a splitter"
Top             =   4425
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "SplitterMove"
Height          =   255
Index           =   10
Left            =   225
TabIndex        =   26
ToolTipText     =   "Occurs when the user is moving a splitter"
Top             =   4050
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "SplitterMouseUp"
Height          =   255
Index           =   9
Left            =   225
TabIndex        =   25
ToolTipText     =   "Occurs when the user releases a mouse button over a splitter without previously moving the splitter"
Top             =   3675
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "SplitterMouseMove"
Height          =   255
Index           =   8
Left            =   225
TabIndex        =   24
ToolTipText     =   "Occurs when the user moves a mouse over a splitter without moving the splitter"
Top             =   3300
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "SplitterMouseDown"
Height          =   255
Index           =   7
Left            =   225
TabIndex        =   23
ToolTipText     =   "Occurs when the user presses a mouse button over a splitter"
Top             =   2925
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "SplitterDblClick"
Height          =   255
Index           =   6
Left            =   225
TabIndex        =   22
ToolTipText     =   "Occurs when the user presses and then realeses a mouse button and then presses and releases it again over a splitter"
Top             =   2550
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "SplitterClick"
Height          =   255
Index           =   5
Left            =   225
TabIndex        =   21
ToolTipText     =   "Occurs when the user presses and then realeses a mouse button over a splitter"
Top             =   2175
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "ControlMoveEnd"
Height          =   255
Index           =   4
Left            =   225
TabIndex        =   20
ToolTipText     =   "Occurs when the user is finished moving a control, i.e. when the rectangle that represents the moving control disappears"
Top             =   1800
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "ControlMoveBegin"
Height          =   255
Index           =   3
Left            =   225
TabIndex        =   19
ToolTipText     =   "Occurs when the user is about to move a control, i.e. the first time the rectangle that represents the moving control occurs"
Top             =   1425
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "ControlMove"
Height          =   255
Index           =   2
Left            =   225
TabIndex        =   18
ToolTipText     =   "Occurs when the user is moving a control"
Top             =   1050
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "ControlBeforeClose"
Height          =   255
Index           =   1
Left            =   225
TabIndex        =   17
ToolTipText     =   "Occurs after the user presses a close button of certain control and before the control is closed"
Top             =   675
Width           =   2580
End
Begin VB.Label lblEvents
Alignment       =   2                             'Center
Caption         =   "ControlAfterClose"
Height          =   255
Index           =   0
Left            =   225
TabIndex        =   16
ToolTipText     =   "Occurs when a control has just been closed by the user"
Top             =   300
Width           =   2580
End
End
Begin VB.Frame fraFeatures
Height          =   7695
Index           =   4
Left            =   390
TabIndex        =   13
Top             =   570
Width           =   3030
Begin VB.Label lblNoCtlMethod
Alignment       =   2                             'Center
Caption         =   "No Method"
Height          =   210
Left            =   45
TabIndex        =   14
Top             =   3765
Width           =   2925
End
End
Begin VB.Frame fraFeatures
Height          =   7695
Index           =   8
Left            =   360
TabIndex        =   11
Top             =   300
Width           =   3030
Begin VB.Label lblNoSplEvent
Alignment       =   2                             'Center
Caption         =   "No Event"
Height          =   210
Left            =   30
TabIndex        =   12
Top             =   3480
Width           =   2925
End
End
Begin VB.Frame fraFeatures
Height          =   7695
Index           =   7
Left            =   390
TabIndex        =   9
Top             =   390
Width           =   3030
Begin VB.Label lblNoSplMethod
Alignment       =   2                             'Center
Caption         =   "No Method"
Height          =   210
Left            =   45
TabIndex        =   10
Top             =   3570
Width           =   2925
End
End
Begin VB.Frame fraFeatures
Height          =   7035
Index           =   5
Left            =   390
TabIndex        =   7
Top             =   300
Width           =   3030
Begin VB.Label lblNoCtlEvent
Alignment       =   2                             'Center
Caption         =   "No Event"
Height          =   255
Left            =   45
TabIndex        =   8
Top             =   3330
Width           =   2895
End
End
Begin MSComctlLib.TabStrip tabFeatures
Height          =   330
Left            =   360
TabIndex        =   5
Top             =   30
Width           =   3090
_ExtentX        =   5450
_ExtentY        =   582
TabWidthStyle   =   1
MultiRow        =   -1                            'True
Style           =   2
HotTracking     =   -1                            'True
Separators      =   -1                            'True
_Version        =   393216
BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628}
NumTabs         =   3
BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628}
Caption         =   "Properties"
ImageVarType    =   2
EndProperty
BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628}
Caption         =   "Methods"
ImageVarType    =   2
EndProperty
BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628}
Caption         =   "Events"
ImageVarType    =   2
EndProperty
EndProperty
BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
Name            =   "MS Sans Serif"
Size            =   8.25
Charset         =   0
Weight          =   700
Underline       =   0                             'False
Italic          =   0                             'False
Strikethrough   =   0                             'False
EndProperty
End
Begin MSComctlLib.TabStrip tabMembers
Height          =   7455
Left            =   0
TabIndex        =   6
Top             =   390
Width           =   435
_ExtentX        =   767
_ExtentY        =   13150
TabWidthStyle   =   1
MultiRow        =   -1                            'True
HotTracking     =   -1                            'True
Placement       =   2
Separators      =   -1                            'True
_Version        =   393216
BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628}
NumTabs         =   3
BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628}
Caption         =   "Control Manager"
ImageVarType    =   2
EndProperty
BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628}
Caption         =   "Control"
ImageVarType    =   2
EndProperty
BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628}
Caption         =   "Splitter"
ImageVarType    =   2
EndProperty
EndProperty
BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
Name            =   "MS Sans Serif"
Size            =   8.25
Charset         =   0
Weight          =   700
Underline       =   0                             'False
Italic          =   0                             'False
Strikethrough   =   0                             'False
EndProperty
End
Begin MSComDlg.CommonDialog cdlColor
Left            =   30
Top             =   60
_ExtentX        =   847
_ExtentY        =   847
_Version        =   393216
End
End
Begin MSComctlLib.TreeView TreeView1
Height          =   1809
Left            =   3647
TabIndex        =   3
Top             =   5878
Width           =   4588
_ExtentX        =   8096
_ExtentY        =   3201
_Version        =   393217
LineStyle       =   1
Style           =   7
Checkboxes      =   -1                            'True
HotTracking     =   -1                            'True
Appearance      =   1
End
Begin MSComctlLib.ListView ListView1
Height          =   1118
Left            =   0
TabIndex        =   2
Top             =   7972
Width           =   8235
_ExtentX        =   14526
_ExtentY        =   1984
View            =   3
MultiSelect     =   -1                            'True
LabelWrap       =   -1                            'True
HideSelection   =   -1                            'True
AllowReorder    =   -1                            'True
Checkboxes      =   -1                            'True
FlatScrollBar   =   -1                            'True
FullRowSelect   =   -1                            'True
GridLines       =   -1                            'True
_Version        =   393217
ForeColor       =   -2147483640
BackColor       =   -2147483643
BorderStyle     =   1
Appearance      =   1
NumItems        =   5
BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628}
Text            =   "Column 1"
Object.Width           =   2540
EndProperty
BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628}
SubItemIndex    =   1
Text            =   "Column 2"
Object.Width           =   2540
EndProperty
BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628}
SubItemIndex    =   2
Text            =   "Column 3"
Object.Width           =   2540
EndProperty
BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628}
SubItemIndex    =   3
Text            =   "Column 4"
Object.Width           =   2540
EndProperty
BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628}
SubItemIndex    =   4
Text            =   "Column 5"
Object.Width           =   2540
EndProperty
End
Begin VB.TextBox Text1
Height          =   2094
Left            =   3647
TabIndex        =   1
Text            =   "TextBox Sample"
Top             =   3499
Width           =   4588
End
End
Begin VB.Timer tmrEvents
Enabled         =   0                             'False
Index           =   0
Interval        =   1000
Left            =   -60
Top             =   6000
End
End
Attribute VB_Name = "frmDemoFeatures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const conFraControlManagerProperties As Long = 0
Private Const conFraControlManagerMethods As Long = 1
Private Const conFraControlManagerEvents As Long = 2
Private Const conFraControlProperties  As Long = 3
Private Const conFraControlMethods     As Long = 4
Private Const conFraControlEvents      As Long = 5
Private Const conFraSplitterProperties As Long = 6
Private Const conFraSplitterMethods    As Long = 7
Private Const conFraSplitterEvents     As Long = 8
Private Const conLblEventControlAfterClose As Long = 0
Private Const conLblEventControlBeforeClose As Long = 1
Private Const conLblEventControlMove   As Long = 2
Private Const conLblEventControlMoveBegin As Long = 3
Private Const conLblEventControlMoveEnd As Long = 4
Private Const conLblEventSplitterClick As Long = 5
Private Const conLblEventSplitterDblClick As Long = 6
Private Const conLblEventSplitterMouseDown As Long = 7
Private Const conLblEventSplitterMouseMove As Long = 8
Private Const conLblEventSplitterMouseUp As Long = 9
Private Const conLblEventSplitterMove  As Long = 10
Private Const conLblEventSplitterMoveBegin As Long = 11
Private Const conLblEventSplitterMoveEnd As Long = 12
Private Const conLblEventTitleBarClick As Long = 13
Private Const conLblEventTitleBarDblClick As Long = 14
Private Const conLblEventTitleBarMouseDown As Long = 15
Private Const conLblEventTitleBarMouseMove As Long = 16
Private Const conLblEventTitleBarMouseUp As Long = 17
Private Const conTabProperties         As Long = 1
Private Const conTabMethods            As Long = 2
Private Const conTabEvents             As Long = 3
Private Const conTabControlManager     As Long = 1
Private Const conTabControl            As Long = 2
Private Const conTabSplitter           As Long = 3
Private lngSelectedFrame               As Long
'*****************************************************
Private Sub cboCMClipCursor_Click()
'*****************************************************
10   ControlManager1.ClipCursor = CBool(cboCMClipCursor)
20   RefreshProperties
End Sub
'*****************************************************
Private Sub cboCMEnable_Click()
'*****************************************************
10   ControlManager1.Enable = CBool(cboCMEnable)
20   RefreshProperties
End Sub
'*****************************************************
Private Sub cboCMFillContainer_Click()
'*****************************************************
10   ControlManager1.FillContainer = CBool(cboCMFillContainer)
End Sub
'*****************************************************
Private Sub cboCMLiveUpdate_Click()
'*****************************************************
10   ControlManager1.LiveUpdate = CBool(cboCMLiveUpdate)
20   RefreshProperties
End Sub
'*****************************************************
Private Sub cboCMTitleBarCloseVisible_Click()
'*****************************************************
10   ControlManager1.TitleBar_CloseVisible = CBool(cboCMTitleBarCloseVisible)
20   RefreshProperties
End Sub
'*****************************************************
Private Sub cboCMTitleBarVisible_Click()
'*****************************************************
10   ControlManager1.TitleBar_Visible = CBool(cboCMTitleBarVisible)
20   RefreshProperties
End Sub
'*****************************************************
Private Sub cboCtlTitleBarCloseVisible_Click()
'*****************************************************
10   If LenB(txtCtlId) Then ControlManager1.Controls(CLng(txtCtlId)).TitleBar_CloseVisible = CBool(cboCtlTitleBarCloseVisible)
End Sub
'*****************************************************
Private Sub cboCtlTitleBarVisible_Click()
'*****************************************************
10   If LenB(txtCtlId) Then
20      ControlManager1.Controls(CLng(txtCtlId)).TitleBar_Visible = CBool(cboCtlTitleBarVisible)
30      RefreshProperties
40      End If
End Sub
'*****************************************************
Private Sub cboSplClipCursor_Click()
'*****************************************************
10   If LenB(txtSplId) Then
20      If ControlManager1.Splitters.IsExist(CLng(txtSplId)) Then ControlManager1.Splitters(CLng(txtSplId)).ClipCursor = CBool(cboSplClipCursor)
30      End If
End Sub
'*****************************************************
Private Sub cboSplEnable_Click()
'*****************************************************
10   If LenB(txtSplId) Then
20      If ControlManager1.Splitters.IsExist(CLng(txtSplId)) Then ControlManager1.Splitters(CLng(txtSplId)).Enable = CBool(cboSplEnable)
30      End If
End Sub
'*****************************************************
Private Sub cboSplLiveUpdate_Click()
'*****************************************************
10   If LenB(txtSplId) Then
20      If ControlManager1.Splitters.IsExist(CLng(txtSplId)) Then ControlManager1.Splitters(CLng(txtSplId)).LiveUpdate = CBool(cboSplLiveUpdate)
30      End If
End Sub
'*****************************************************
Private Sub ChangeTab()
'*****************************************************
10   Select Case tabMembers.SelectedItem.Index
        Case conTabControlManager
20         Select Case tabFeatures.SelectedItem.Index
              Case conTabProperties
30               lngSelectedFrame = conFraControlManagerProperties
40            Case conTabMethods
50               lngSelectedFrame = conFraControlManagerMethods
60            Case conTabEvents
70               lngSelectedFrame = conFraControlManagerEvents
80            End Select
90      Case conTabControl
100         Select Case tabFeatures.SelectedItem.Index
               Case conTabProperties
110               lngSelectedFrame = conFraControlProperties
120            Case conTabMethods
130               lngSelectedFrame = conFraControlMethods
140            Case conTabEvents
150               lngSelectedFrame = conFraControlEvents
160            End Select
170      Case conTabSplitter
180         Select Case tabFeatures.SelectedItem.Index
               Case conTabProperties
190               lngSelectedFrame = conFraSplitterProperties
200            Case conTabMethods
210               lngSelectedFrame = conFraSplitterMethods
220            Case conTabEvents
230               lngSelectedFrame = conFraSplitterEvents
240            End Select
250      End Select
260   ShowSelectedFrame
End Sub
'*****************************************************
Private Sub ClearEvent(ByVal lngId As Long)
'*****************************************************
10   lblEvents(lngId).BorderStyle = vbBSNone
20   lblEvents(lngId).Font.Bold = False
30   tmrEvents(lngId).Enabled = False
End Sub
'*****************************************************
Private Sub cmdCloseControl_Click()
'*****************************************************
   Dim blnSuccess          As Boolean
10   On Error GoTo ErrorHandler
20   ControlManager1.ShowControl CLng(lblIdControl(0)), False, blnSuccess
30   If Not blnSuccess Then ShowErrMessage "Fail to close the control"
40   Exit Sub
50   ErrorHandler:
60   ShowErrMessage
End Sub
'*****************************************************
Private Sub cmdMoveControl_Click()
'*****************************************************
   Dim blnSuccess          As Boolean
10   On Error GoTo ErrorHandler
20   Select Case cboMoveControlMoveTo.ListIndex
        Case mdSplitter
30         blnSuccess = ControlManager1.MoveControl(lblIdControl(1), cboMoveControlMoveTo.ListIndex, IdSplitterDestination:=CLng(lblIdSplitter(0)))
40      Case mdControlTop, mdControlRight, mdControlBottom, mdControlLeft
50         blnSuccess = ControlManager1.MoveControl(lblIdControl(1), cboMoveControlMoveTo.ListIndex, lblIdControl2)
60      Case mdEdgeTop, mdEdgeRight, mdEdgeBottom, mdEdgeLeft
70         blnSuccess = ControlManager1.MoveControl(lblIdControl(1), cboMoveControlMoveTo.ListIndex)
80      End Select
90   If Not blnSuccess Then ShowErrMessage "Fail to move the control"
100   Exit Sub
110   ErrorHandler:
120   ShowErrMessage
End Sub
'*****************************************************
Private Sub cmdMoveSplitter_Click()
'*****************************************************
10   On Error GoTo ErrorHandler
20   ControlManager1.MoveSplitter IdSplitter:=CLng(lblIdSplitter(1)), MoveTo:=CLng(txtMoveSplitterMoveTo)
30   Exit Sub
40   ErrorHandler:
50   ShowErrMessage
End Sub
'*****************************************************
Private Sub cmdOpenControl_Click()
'*****************************************************
   Dim blnSuccess          As Boolean
10   On Error GoTo ErrorHandler
20   ControlManager1.ShowControl lblIdControl(2), True, blnSuccess, CBool(cboOpenControlMaintainSize)
30   If Not blnSuccess Then ShowErrMessage "Fail to open the control"
40   Exit Sub
50   ErrorHandler:
60   ShowErrMessage
End Sub
'*****************************************************
Private Sub ControlManager1_ControlAfterClose(ByVal sIdControl As String)
'*****************************************************
10   HighlightEvent conLblEventControlAfterClose
20   RefreshProperties
End Sub
'*****************************************************
Private Sub ControlManager1_ControlBeforeClose(ByVal sIdControl As String, Cancel As Boolean)
'*****************************************************
10   HighlightEvent conLblEventControlBeforeClose
End Sub
'*****************************************************
Private Sub ControlManager1_ControlMove(ByVal sIdControl As String, ByVal Shift As Integer, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
'*****************************************************
10   HighlightEvent conLblEventControlMove
End Sub
'*****************************************************
Private Sub ControlManager1_ControlMoveBegin(ByVal sIdControl As String, ByVal Shift As Integer)
'*****************************************************
10   HighlightEvent conLblEventControlMoveBegin
End Sub
'*****************************************************
Private Sub ControlManager1_ControlMoveEnd(ByVal sIdControl As String, ByVal Shift As Integer, ByVal Moved As Boolean)
'*****************************************************
10   HighlightEvent conLblEventControlMoveEnd
20   RefreshProperties
End Sub
'*****************************************************
Private Sub ControlManager1_SplitterClick(ByVal IdSplitter As Long)
'*****************************************************
   Dim i                   As Long
10   HighlightEvent conLblEventSplitterClick
20   If fraFeatures(conFraControlManagerMethods).Visible Then
30      For i = 0 To lblIdSplitter.UBound
40         lblIdSplitter(i) = CStr(IdSplitter)
50         Next
60      End If
70   txtSplId = CStr(IdSplitter)
End Sub
'*****************************************************
Private Sub ControlManager1_SplitterDblClick(ByVal IdSplitter As Long)
'*****************************************************
10   HighlightEvent conLblEventSplitterDblClick
End Sub
'*****************************************************
Private Sub ControlManager1_SplitterMouseDown(ByVal IdSplitter As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'*****************************************************
10   HighlightEvent conLblEventSplitterMouseDown
End Sub
'*****************************************************
Private Sub ControlManager1_SplitterMouseMove(ByVal IdSplitter As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'*****************************************************
10   HighlightEvent conLblEventSplitterMouseMove
End Sub
'*****************************************************
Private Sub ControlManager1_SplitterMouseUp(ByVal IdSplitter As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'*****************************************************
10   HighlightEvent conLblEventSplitterMouseUp
End Sub
'*****************************************************
Private Sub ControlManager1_SplitterMove(ByVal IdSplitter As Long, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'*****************************************************
10   HighlightEvent conLblEventSplitterMove
End Sub
'*****************************************************
Private Sub ControlManager1_SplitterMoveBegin(ByVal IdSplitter As Long, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'*****************************************************
10   HighlightEvent conLblEventSplitterMoveBegin
End Sub
'*****************************************************
Private Sub ControlManager1_SplitterMoveEnd(ByVal IdSplitter As Long, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'*****************************************************
10   HighlightEvent conLblEventSplitterMoveEnd
20   RefreshProperties
End Sub
'*****************************************************
Private Sub ControlManager1_TitleBarClick(ByVal sIdControl As String)
'*****************************************************
   Dim i                   As Long
10   HighlightEvent conLblEventTitleBarClick
20   If fraFeatures(conFraControlManagerMethods).Visible Then
30      For i = 0 To lblIdControl.UBound
40         lblIdControl(i) = sIdControl
50         Next
60      End If
70   txtCtlId = sIdControl
End Sub
'*****************************************************
Private Sub ControlManager1_TitleBarDblClick(ByVal sIdControl As String)
'*****************************************************
10   HighlightEvent conLblEventTitleBarDblClick
20   If fraFeatures(conFraControlManagerMethods).Visible Then
30      lblIdControl2 = sIdControl
40      End If
End Sub
'*****************************************************
Private Sub ControlManager1_TitleBarMouseDown(ByVal sIdControl As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'*****************************************************
10   HighlightEvent conLblEventTitleBarMouseDown
End Sub
'*****************************************************
Private Sub ControlManager1_TitleBarMouseMove(ByVal sIdControl As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'*****************************************************
10   HighlightEvent conLblEventTitleBarMouseMove
End Sub
'*****************************************************
Private Sub ControlManager1_TitleBarMouseUp(ByVal sIdControl As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'*****************************************************
10   HighlightEvent conLblEventTitleBarMouseUp
End Sub
'*****************************************************
Private Sub Form_Load()
'*****************************************************
   Dim i                   As Long
   Dim Index               As Long
   Dim j                   As Long
   Dim k                   As Long
   Dim nodi                As Node
   Dim nodj                As Node
10   Const conFrameTop As Long = 300&
20   Const conFrameLeft As Long = 390&
30   Const conFrameHeight As Long = 7690&
40   Const conFrameWidth As Long = 3030&
50   On Error Resume Next                         'incase docs were moved/deleted
60   WBFeatures.Navigate App.Path & "\docs\TheoZFeatures.htm"
70   DoEvents                                     'MUST GIVE TIME TO COMPLETE TO AVOID IE7 Runtime Error 7
80   On Error GoTo 0
90   For i = lblEvents.lBound + 1 To lblEvents.UBound
100      Load tmrEvents(i)
110      Next
120   For i = 0 To 8
130      With fraFeatures(i)
140         .Move conFrameLeft, conFrameTop, conFrameWidth, conFrameHeight
150         End With
160      Next
170   With Me.fraConCtlProperties
180      .Move 0, 0, conFrameWidth, conFrameHeight
190      End With
200   With Me.fraConSplProperties
210      .Move 0, 0, conFrameWidth, conFrameHeight
220      End With
230   With Me.fraCtlProperties
240      .Move 0, conFrameTop, conFrameWidth, conFrameHeight
250      End With
260   With Me.fraSplProperties
270      .Move 0, conFrameTop, conFrameWidth, conFrameHeight
280      End With
290   For i = 1 To 10
300      Set nodi = TreeView1.Nodes.Add(, , , "Node " & CStr(i))
310      For j = 1 To 6
320         Set nodj = TreeView1.Nodes.Add(nodi.Index, tvwChild, , "Node " & CStr(i) & "." & CStr(j))
330         For k = 1 To 3
340            TreeView1.Nodes.Add nodj.Index, tvwChild, , "Node " & CStr(i) & "." & CStr(j) & "." & CStr(k)
350            Next
360         Next
370      Next
380   For i = 1 To 10
390      ListView1.ListItems.Add , , "Item " & CStr(i) & ".1"
400      For j = 1 To ListView1.ColumnHeaders.Count - 1
410         ListView1.ListItems(i).SubItems(j) = "Item " & CStr(i) & "." & CStr(j + 1)
420         Next
430      Next
440   tabFeatures_Click
450   InitFeatures
End Sub
'*****************************************************
Private Sub HighlightEvent(ByVal lngId As Long)
'*****************************************************
10   If fraFeatures(conFraControlManagerEvents).Visible Then
20      lblEvents(lngId).Font.Bold = True
30      lblEvents(lngId).BorderStyle = vbFixedSingle
40      tmrEvents(lngId).Enabled = True
50      End If
End Sub
'*****************************************************
Private Sub InitFeatures()
'*****************************************************
10   fraConCtlProperties.Height = txtCtlYc.Top + txtCtlYc.Height + (7 * Screen.TwipsPerPixelY)
20   fraCtlProperties.Height = fraConCtlProperties.Height
30   fraConSplProperties.Height = txtSplYc.Top + txtSplYc.Height + (7 * Screen.TwipsPerPixelY)
40   fraSplProperties.Height = fraConSplProperties.Height
50   With ControlManager1
60      lblCMActiveColor.BackColor = .ActiveColor
70      lblCMBackColor.BackColor = .BackColor
80      cboCMClipCursor = CStr(.ClipCursor)
90      cboCMEnable = CStr(.Enable)
100      cboCMFillContainer = CStr(.FillContainer)
110      cboCMLiveUpdate = CStr(.LiveUpdate)
120      cboCMTitleBarCloseVisible = CStr(.TitleBar_CloseVisible)
130      cboCMTitleBarVisible = CStr(.TitleBar_Visible)
140      txtCMMarginBottom = CStr(.MarginBottom)
150      txtCMMarginLeft = CStr(.MarginLeft)
160      txtCMMarginRight = CStr(.MarginRight)
170      txtCMMarginTop = CStr(.MarginTop)
180      txtCMTitleBarHeight = CStr(.TitleBar_Height)
190      txtCMSize = CStr(.Size)
200      End With
210   cboMoveControlMoveTo = "mdSplitter"
220   cboOpenControlMaintainSize = "True"
End Sub
'*****************************************************
Private Sub lblCMActiveColor_Click()
'*****************************************************
10   cdlColor.Flags = cdlCCRGBInit
20   cdlColor.Color = lblCMActiveColor.BackColor
30   cdlColor.ShowColor
40   lblCMActiveColor.BackColor = cdlColor.Color
50   ControlManager1.ActiveColor = lblCMActiveColor.BackColor
End Sub
'*****************************************************
Private Sub lblCMBackColor_Click()
'*****************************************************
10   cdlColor.Flags = cdlCCRGBInit
20   cdlColor.Color = lblCMBackColor.BackColor
30   cdlColor.ShowColor
40   lblCMBackColor.BackColor = cdlColor.Color
50   ControlManager1.BackColor = lblCMBackColor.BackColor
End Sub
'*****************************************************
Private Sub lblCtlClick_Click()
'*****************************************************
10   MsgBox "Click a title bar to see its control properties", vbInformation
End Sub
'*****************************************************
Private Sub lblSplActiveColor_Click()
'*****************************************************
10   If LenB(txtSplId) Then
20      If ControlManager1.Splitters.IsExist(CLng(txtSplId)) Then
30         cdlColor.Flags = cdlCCRGBInit
40         cdlColor.Color = lblSplActiveColor.BackColor
50         cdlColor.ShowColor
60         lblSplActiveColor.BackColor = cdlColor.Color
70         ControlManager1.Splitters(CLng(txtSplId)).ActiveColor = lblSplActiveColor.BackColor
80         End If
90      End If
End Sub
'*****************************************************
Private Sub lblSplBackColor_Click()
'*****************************************************
10   If LenB(txtSplId) Then
20      If ControlManager1.Splitters.IsExist(CLng(txtSplId)) Then
30         cdlColor.Flags = cdlCCRGBInit
40         cdlColor.Color = lblSplBackColor.BackColor
50         cdlColor.ShowColor
60         lblSplBackColor.BackColor = cdlColor.Color
70         ControlManager1.Splitters(CLng(txtSplId)).BackColor = lblSplBackColor.BackColor
80         End If
90      End If
End Sub
'*****************************************************
Private Sub lblSplClick_Click()
'*****************************************************
10   MsgBox "Click a splitter to see its properties", vbInformation
End Sub
'*****************************************************
Private Sub picMain_Resize()
'*****************************************************
   Dim lngNewHeight        As Long
   Dim TwipsPerPixelY11    As Long
   Dim TwipsPerPixelY3     As Long
10   On Error GoTo picMain_Resize_Err
     'Dim lngNewHeight As Long, TwipsPerPixelY11 As Long, TwipsPerPixelY3 As Long
20   TwipsPerPixelY3 = 3 * Screen.TwipsPerPixelY
30   TwipsPerPixelY11 = 11 * Screen.TwipsPerPixelY
40   lngNewHeight = picMain.Height - (tabFeatures.Top + tabFeatures.Height)
50   vsbCtlProperties.Visible = (txtCtlYc.Top + txtCtlYc.Height + TwipsPerPixelY3 > lngNewHeight)
60   vsbSplProperties.Visible = (txtSplYc.Top + txtSplYc.Height + TwipsPerPixelY3 > lngNewHeight)
70   If lngNewHeight - TwipsPerPixelY11 > 0 Then
80      fraFeatures(lngSelectedFrame).Height = lngNewHeight
90      tabMembers.Height = lngNewHeight - (7 * Screen.TwipsPerPixelY)
100      vsbCtlProperties.Height = lngNewHeight - (8 * Screen.TwipsPerPixelY)
110      vsbSplProperties.Height = vsbCtlProperties.Height
120      fraConCtlProperties.Height = lngNewHeight - TwipsPerPixelY11
130      fraConSplProperties.Height = fraConCtlProperties.Height
140      lblNoCtlMethod.Top = (lngNewHeight \ 2) - (lblNoCtlMethod.Height \ 4)
150      lblNoCtlEvent.Top = lblNoCtlMethod.Top
160      lblNoSplMethod.Top = lblNoCtlMethod.Top
170      lblNoSplEvent.Top = lblNoCtlMethod.Top
180   Else
190      fraFeatures(lngSelectedFrame).Height = 0
200      tabMembers.Height = 0
210      vsbCtlProperties.Height = 0
220      vsbSplProperties.Height = 0
230      fraConCtlProperties.Height = 0
240      fraConSplProperties.Height = 0
250      lblNoCtlMethod.Top = 0
260      lblNoCtlEvent.Top = 0
270      End If
280   RefreshScrollBar
290   picMain_Resize_Exit:
300   On Error GoTo 0
310   Exit Sub
320   picMain_Resize_Err:
330   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", picMain_Resize", Name
340   Resume picMain_Resize_Exit
End Sub
'*****************************************************
Private Sub RefreshProperties()
'*****************************************************
   Static blnRefreshing    As Boolean
   Dim oid                 As clsid
10   On Error GoTo RefreshProperties_Err
20   If Not blnRefreshing Then
30      blnRefreshing = True
40      With ControlManager1
50         txtCMTitleBarHeight = CStr(.TitleBar_Height)
60         cboCMTitleBarVisible = CStr(.TitleBar_Visible)
70         End With
80      If LenB(txtCtlId) Then
90         With ControlManager1.Controls(txtCtlId)
100            txtCtlBottom = CStr(.bottom)
110            cboCtlClosed = CStr(.Closed)
120            txtCtlHeight = CStr(.Height)
130            txtCtlIdSplBottom = CStr(.IdSplBottom)
140            txtCtlIdSplLeft = CStr(.IdSplLeft)
150            txtCtlIdSplRight = CStr(.IdSplRight)
160            txtCtlIdSplTop = CStr(.IdSplTop)
170            txtCtlLeft = CStr(.Left)
180            txtCtlMinHeight = CStr(.MinHeight)
190            txtCtlMinWidth = CStr(.MinWidth)
200            txtCtlName = .Key
210            txtCtlRight = CStr(.Right)
220            cboCtlTitleBarCloseVisible = CStr(.TitleBar_CloseVisible)
230            txtCtlTitleBarHeight = CStr(.TitleBar_Height)
240            cboCtlTitleBarVisible = CStr(.TitleBar_Visible)
250            txtCtlTop = CStr(.Top)
260            txtCtlWidth = CStr(.Width)
270            txtCtlXc = CStr(.Xc)
280            txtCtlYc = CStr(.Yc)
290            End With
300         End If
310      If LenB(txtSplId) Then
320         If Not ControlManager1.Splitters.IsExist(CLng(txtSplId)) Then
330            lblSplActiveColor.BackColor = vbBlack
340            lblSplBackColor.BackColor = vbBlack
350            txtSplBottom = vbNullString
360            cboSplClipCursor.ListIndex = -1
370            cboSplEnable.ListIndex = -1
380            txtSplHeight = vbNullString
390            txtSplId = vbNullString
400            lstSplIdsCtlBottom.Clear
410            lstSplIdsCtlLeft.Clear
420            lstSplIdsCtlRight.Clear
430            lstSplIdsCtlTop.Clear
440            lstSplIdsSplBottom.Clear
450            lstSplIdsSplLeft.Clear
460            lstSplIdsSplRight.Clear
470            lstSplIdsSplTop.Clear
480            txtSplLeft = vbNullString
490            cboSplLiveUpdate.ListIndex = -1
500            txtSplMaxXc = vbNullString
510            txtSplMaxYc = vbNullString
520            txtSplMinXc = vbNullString
530            txtSplMinYc = vbNullString
540            cboSplOrientation.ListIndex = -1
550            txtSplRight = vbNullString
560            txtSplTop = vbNullString
570            txtSplWidth = vbNullString
580            txtSplXc = vbNullString
590            txtSplYc = vbNullString
600         Else
610            With ControlManager1.Splitters(CLng(txtSplId))
620               lblSplActiveColor.BackColor = .ActiveColor
630               lblSplBackColor.BackColor = .BackColor
640               txtSplBottom = CStr(.bottom)
650               cboSplClipCursor = CStr(.ClipCursor)
660               cboSplEnable = CStr(.Enable)
670               txtSplHeight = CStr(.Height)
680               lstSplIdsCtlBottom.Clear
690               For Each oid In .IdsCtlBottom
700                  lstSplIdsCtlBottom.AddItem CStr(oid)
710                  Next
720               lstSplIdsCtlLeft.Clear
730               For Each oid In .IdsCtlLeft
740                  lstSplIdsCtlLeft.AddItem CStr(oid)
750                  Next
760               lstSplIdsCtlRight.Clear
770               For Each oid In .IdsCtlRight
780                  lstSplIdsCtlRight.AddItem CStr(oid)
790                  Next
800               lstSplIdsCtlTop.Clear
810               For Each oid In .IdsCtlTop
820                  lstSplIdsCtlTop.AddItem CStr(oid)
830                  Next
840               lstSplIdsSplBottom.Clear
850               For Each oid In .IdsSplBottom
860                  lstSplIdsSplBottom.AddItem CStr(oid)
870                  Next
880               lstSplIdsSplLeft.Clear
890               For Each oid In .IdsSplLeft
900                  lstSplIdsSplLeft.AddItem CStr(oid)
910                  Next
920               lstSplIdsSplRight.Clear
930               For Each oid In .IdsSplRight
940                  lstSplIdsSplRight.AddItem CStr(oid)
950                  Next
960               lstSplIdsSplTop.Clear
970               For Each oid In .IdsSplTop
980                  lstSplIdsSplTop.AddItem CStr(oid)
990                  Next
1000               txtSplLeft = CStr(.Left)
1010               cboSplLiveUpdate = CStr(.LiveUpdate)
1020               txtSplMaxXc = CStr(.MaxXc)
1030               txtSplMaxYc = CStr(.MaxYc)
1040               txtSplMinXc = CStr(.MinXc)
1050               txtSplMinYc = CStr(.MinYc)
1060               Select Case .Orientation
                      Case orHorizontal
1070                     cboSplOrientation = "orHorizontal"
1080                  Case orVertical
1090                     cboSplOrientation = "orVertical"
1100                  End Select
1110               txtSplRight = CStr(.Right)
1120               txtSplTop = CStr(.Top)
1130               txtSplWidth = CStr(.Width)
1140               txtSplXc = CStr(.Xc)
1150               txtSplYc = CStr(.Yc)
1160               End With
1170            End If
1180         End If
1190      blnRefreshing = False
1200      End If
1210   RefreshProperties_Exit:
1220   On Error GoTo 0
1230   Exit Sub
1240   RefreshProperties_Err:
1250   ErrHandler Err, Error$, "Line:" & VBA.Erl & ", RefreshProperties", Name
1260   Resume RefreshProperties_Exit
End Sub
'*****************************************************
Private Sub RefreshScrollBar()
'*****************************************************
10   With vsbCtlProperties
20      .Min = 0
30      .Max = fraCtlProperties.Height - fraFeatures(conFraControlProperties).Height + (7 * Screen.TwipsPerPixelY)
40      .SmallChange = Screen.TwipsPerPixelY * 10
50      .LargeChange = Screen.TwipsPerPixelY * 100
60      End With
70   With vsbSplProperties
80      .Min = 0
90      .Max = fraSplProperties.Height - fraFeatures(conFraSplitterProperties).Height + (7 * Screen.TwipsPerPixelY)
100      .SmallChange = Screen.TwipsPerPixelY * 10
110      .LargeChange = Screen.TwipsPerPixelY * 100
120      End With
End Sub
'*****************************************************
Private Sub ShowErrMessage(Optional strErrMessage As String = vbNullString)
'*****************************************************
10   If strErrMessage = vbNullString Then strErrMessage = Err.Description
20   MsgBox strErrMessage, vbCritical + vbOKOnly
End Sub
'*****************************************************
Private Sub ShowSelectedFrame()
'*****************************************************
   Dim i                   As Integer
10   For i = fraFeatures.lBound To fraFeatures.UBound
20      fraFeatures(i).Visible = False
30      Next
40   With fraFeatures(lngSelectedFrame)
50      .Visible = True
60      .ZOrder
70      End With
80   picMain_Resize
End Sub
'*****************************************************
Private Sub tabFeatures_Click()
'*****************************************************
10   ChangeTab
End Sub
'*****************************************************
Private Sub tabMembers_Click()
'*****************************************************
10   ChangeTab
End Sub
'*****************************************************
Private Sub tmrEvents_Timer(Index As Integer)
'*****************************************************
10   ClearEvent Index
End Sub
'*****************************************************
Private Sub txtCMMarginBottom_Validate(Cancel As Boolean)
'*****************************************************
   Dim lngOldValue         As Long
10   On Error GoTo ErrorHandler
20   lngOldValue = ControlManager1.MarginBottom
30   ControlManager1.MarginBottom = CLng(txtCMMarginBottom)
40   Exit Sub
50   ErrorHandler:
60   ShowErrMessage
70   txtCMMarginBottom = lngOldValue
80   Cancel = True
End Sub
'*****************************************************
Private Sub txtCMMarginLeft_Validate(Cancel As Boolean)
'*****************************************************
   Dim lngOldValue         As Long
10   On Error GoTo ErrorHandler
20   lngOldValue = ControlManager1.MarginLeft
30   ControlManager1.MarginLeft = CLng(txtCMMarginLeft)
40   Exit Sub
50   ErrorHandler:
60   ShowErrMessage
70   txtCMMarginLeft = lngOldValue
80   Cancel = True
End Sub
'*****************************************************
Private Sub txtCMMarginRight_Validate(Cancel As Boolean)
'*****************************************************
   Dim lngOldValue         As Long
10   On Error GoTo ErrorHandler
20   lngOldValue = ControlManager1.MarginRight
30   ControlManager1.MarginRight = CLng(txtCMMarginRight)
40   Exit Sub
50   ErrorHandler:
60   ShowErrMessage
70   txtCMMarginRight = lngOldValue
80   Cancel = True
End Sub
'*****************************************************
Private Sub txtCMMarginTop_Validate(Cancel As Boolean)
'*****************************************************
   Dim lngOldValue         As Long
10   On Error GoTo ErrorHandler
20   lngOldValue = ControlManager1.MarginTop
30   ControlManager1.MarginTop = CLng(txtCMMarginTop)
40   Exit Sub
50   ErrorHandler:
60   ShowErrMessage
70   txtCMMarginTop = lngOldValue
80   Cancel = True
End Sub
'*****************************************************
Private Sub txtCMSize_Validate(Cancel As Boolean)
'*****************************************************
   Dim lngOldValue         As Long
10   On Error GoTo ErrorHandler
20   lngOldValue = ControlManager1.Size
30   ControlManager1.Size = CLng(txtCMSize)
40   txtCMSize = CStr(ControlManager1.Size)
50   Exit Sub
60   ErrorHandler:
70   ShowErrMessage
80   txtCMSize = lngOldValue
90   ControlManager1.Size = lngOldValue
100   Cancel = True
End Sub
'*****************************************************
Private Sub txtCtlId_Change()
'*****************************************************
10   RefreshProperties
End Sub
'*****************************************************
Private Sub txtSplId_Change()
'*****************************************************
10   If LenB(txtSplId) Then RefreshProperties
End Sub
'*****************************************************
Private Sub vsbCtlProperties_Change()
'*****************************************************
10   fraCtlProperties.Top = -vsbCtlProperties.Value
End Sub
'*****************************************************
Private Sub vsbSplProperties_Change()
'*****************************************************
10   fraSplProperties.Top = -vsbSplProperties.Value
End Sub

' Yorgi's 4Matz [Feb 28,2007 23:58:48] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

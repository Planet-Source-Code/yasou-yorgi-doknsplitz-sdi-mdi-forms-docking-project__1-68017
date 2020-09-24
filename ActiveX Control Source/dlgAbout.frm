VERSION 5.00
Begin VB.Form dlgAbout
BackColor       =   &H00FFFFC0&
BorderStyle     =   3                             'Fixed Dialog
Caption         =   "DoknSplitz - The SDI/MDI Forms Docking Project"
ClientHeight    =   2430
ClientLeft      =   2340
ClientTop       =   1935
ClientWidth     =   5220
ClipControls    =   0                             'False
Icon            =   "dlgAbout.frx":0000
LinkTopic       =   "Form2"
MaxButton       =   0                             'False
MinButton       =   0                             'False
ScaleHeight     =   162
ScaleMode       =   3                             'Pixel
ScaleWidth      =   348
ShowInTaskbar   =   0                             'False
StartUpPosition =   1                             'CenterOwner
Begin VB.Frame Frame1
Height          =   150
Left            =   -60
TabIndex        =   2
Top             =   780
WhatsThisHelpID =   10385
Width           =   5385
End
Begin VB.Label lblEMail
BackStyle       =   0                             'Transparent
Caption         =   "(yorgi@omnisoftsystems.com)"
Height          =   255
Left            =   2220
MouseIcon       =   "dlgAbout.frx":000C
MousePointer    =   99                            'Custom
TabIndex        =   3
Top             =   510
Width           =   2115
End
Begin VB.Label lblName
BackStyle       =   0                             'Transparent
Caption         =   "DoknSplitz"
BeginProperty Font
Name            =   "MS Sans Serif"
Size            =   9.75
Charset         =   0
Weight          =   700
Underline       =   0                             'False
Italic          =   0                             'False
Strikethrough   =   0                             'False
EndProperty
ForeColor       =   &H00000000&
Height          =   285
Left            =   915
TabIndex        =   0
Top             =   180
WhatsThisHelpID =   10382
Width           =   4185
End
Begin VB.Label lblCopyright
BackStyle       =   0                             'Transparent
Caption         =   "Written by Yorgi"
Height          =   225
Left            =   930
TabIndex        =   1
Top             =   510
WhatsThisHelpID =   10383
Width           =   1275
End
Begin VB.Image Image1
Height          =   600
Left            =   135
Picture         =   "dlgAbout.frx":0316
Stretch         =   -1                            'True
Top             =   165
Width           =   630
End
Begin VB.Label lblDesc
BackColor       =   &H00FFFFFF&
BorderStyle     =   1                             'Fixed Single
Caption         =   $"dlgAbout.frx":0940
Height          =   1410
Left            =   30
TabIndex        =   4
Top             =   960
Width           =   5145
End
End
Attribute VB_Name = "dlgAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
Private Sub Form_load()
'*****************************************************
10   lblName = "DoknSplitz v" & App.Major & "." & App.Minor
20   lblDesc = lblDesc & vbNewLine & vbNewLine & "...Special thanks to Theo Zacharias for VB Control Manager" & vbNewLine & "...To Steve McMahon excellent code www.vbaccelerator.com" & vbNewLine & "...and all the great contributions found on PSC && DevX"
End Sub
'*****************************************************
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
10   lblEMail.Font.Underline = False
End Sub
'*****************************************************
Private Sub lblEMail_Click()
'*****************************************************
10   mdlAPI.ShellExecute hWnd, vbNullString, "mailto:yorgi@omnisoftsystems.com", vbNullString, vbNullString, 1
End Sub
'*****************************************************
Private Sub lblEMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
10   lblEMail.Font.Underline = True
End Sub

' Yorgi's 4Matz [Feb 28,2007 23:58:51] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

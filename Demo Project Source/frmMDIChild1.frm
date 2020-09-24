VERSION 5.00
Begin VB.Form frmMDIChild1
BackColor       =   &H00FFFFFF&
Caption         =   "MDIChild1"
ClientHeight    =   1650
ClientLeft      =   60
ClientTop       =   420
ClientWidth     =   3735
LinkTopic       =   "Form1"
MDIChild        =   -1                            'True
ScaleHeight     =   1650
ScaleWidth      =   3735
Begin VB.Label Label1
BackStyle       =   0                             'Transparent
Caption         =   "You can manipulate SDI forms while maintaining full use of your MDI child forms!"
BeginProperty Font
Name            =   "MS Sans Serif"
Size            =   12
Charset         =   0
Weight          =   700
Underline       =   0                             'False
Italic          =   0                             'False
Strikethrough   =   0                             'False
EndProperty
ForeColor       =   &H000000FF&
Height          =   1305
Left            =   120
TabIndex        =   0
Top             =   270
Width           =   3075
End
End
Attribute VB_Name = "frmMDIChild1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
Private Sub Form_Resize()
'*****************************************************
10   Label1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

' Yorgi's 4Matz [Feb 28,2007 23:58:48] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

VERSION 5.00
Begin VB.Form frmWait
BackColor       =   &H000000FF&
BorderStyle     =   3                             'Fixed Dialog
ClientHeight    =   450
ClientLeft      =   45
ClientTop       =   45
ClientWidth     =   3795
ControlBox      =   0                             'False
LinkTopic       =   "Form1"
MaxButton       =   0                             'False
MinButton       =   0                             'False
ScaleHeight     =   450
ScaleWidth      =   3795
ShowInTaskbar   =   0                             'False
StartUpPosition =   2                             'CenterScreen
Begin VB.Label Label1
AutoSize        =   -1                            'True
BackStyle       =   0                             'Transparent
Caption         =   "Loading Demo....Please Wait!"
BeginProperty Font
Name            =   "Tahoma"
Size            =   12
Charset         =   0
Weight          =   700
Underline       =   0                             'False
Italic          =   0                             'False
Strikethrough   =   0                             'False
EndProperty
ForeColor       =   &H00FFFFFF&
Height          =   285
Left            =   150
TabIndex        =   0
Top             =   90
Width           =   3570
End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bReadyToClose                   As Boolean
'*****************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*****************************************************
10   Cancel = bReadyToClose
End Sub
'*****************************************************
Private Sub Form_Resize()
'*****************************************************
10   DoEvents                                     'a lot of work, take a breather
End Sub
'*****************************************************
Public Sub ShutDown()
'*****************************************************
10   bReadyToClose = True
20   Unload Me
End Sub

' Yorgi's 4Matz [Feb 28,2007 23:58:48] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

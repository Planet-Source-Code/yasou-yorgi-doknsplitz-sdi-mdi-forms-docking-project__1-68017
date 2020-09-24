VERSION 5.00
Begin VB.Form frmDonaldDuck
BorderStyle     =   5                             'Sizable ToolWindow
Caption         =   "Donald DoknDuck"
ClientHeight    =   1155
ClientLeft      =   60
ClientTop       =   300
ClientWidth     =   1215
LinkTopic       =   "Form1"
MaxButton       =   0                             'False
MinButton       =   0                             'False
ScaleHeight     =   1155
ScaleWidth      =   1215
ShowInTaskbar   =   0                             'False
StartUpPosition =   3                             'Windows Default
Visible         =   0                             'False
Begin VB.Image Image1
BorderStyle     =   1                             'Fixed Single
Height          =   1050
Left            =   0
Picture         =   "DonaldDuck.frx":0000
Stretch         =   -1                            'True
Top             =   0
Width           =   1155
End
End
Attribute VB_Name = "frmDonaldDuck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
Private Sub Form_Resize()
'*****************************************************
10   Image1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

' Yorgi's 4Matz [Feb 28,2007 23:58:48] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

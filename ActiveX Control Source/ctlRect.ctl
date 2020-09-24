VERSION 5.00
Begin VB.UserControl ctlRect
AutoRedraw      =   -1                            'True
BackStyle       =   0                             'Transparent
ClientHeight    =   1320
ClientLeft      =   0
ClientTop       =   0
ClientWidth     =   1830
DrawWidth       =   3
FillStyle       =   0                             'Solid
ForeColor       =   &H00000000&
ScaleHeight     =   1320
ScaleWidth      =   1830
Begin VB.Shape Shape1
BorderColor     =   &H00008080&
BorderWidth     =   6
Height          =   945
Left            =   120
Top             =   210
Width           =   1605
End
End
Attribute VB_Name = "ctlRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** File Name   : ctlRect.ctl                                                 **
'** Description : A custom rectangle ActiveX control                          **
'** Usage       : As the rectangle that represent the moving control          **
'** Dependencies: mdlGeneral, mdlAPI                                          **
'** Members     :                                                             **
'**   * Collections: -                                                        **
'**   * Objects    : -                                                        **
'**   * Properties : -                                                        **
'**   * Method     : UpdatePosition                                           **
'**   * Events     : -                                                        **
'** Last modified on November 1, 2003                                         **
'*******************************************************************************
Option Explicit
'-- Variables to save the original rectangle size and position
Private mlngRelLeft                    As Long
Private mlngRelTop                     As Long
Private mlngWidth                      As Long
Private mlngHeight                     As Long
'*****************************************************
Public Sub UpdatePosition()
'*****************************************************
   ' Purpose    - Saves the size and position (relative to the current position of
   '              the cursor) of the original rectangle
   ' Effect     - As specified
   ' Purpose    - Update rectangle position based on its original size and position
   '              (relative to the current position of the cursor when the first
   '              time the rectangle is shown)
   Dim uposCursor          As POINTAPI
10   uposCursor = GetCursorRelPos(UserControl.hWnd)
20   Extender.Move Extender.Left + uposCursor.X - mlngRelLeft, Extender.Top + uposCursor.Y - mlngRelTop, mlngWidth, mlngHeight
End Sub
'*****************************************************
Private Sub UserControl_Resize()
'*****************************************************
10   Shape1.Move 0, 0, UserControl.Width, UserControl.Height
End Sub
'*****************************************************
Private Sub UserControl_Show()
'*****************************************************
   Dim uposCursor          As POINTAPI
10   uposCursor = GetCursorRelPos(UserControl.hWnd)
20   mlngRelLeft = uposCursor.X
30   mlngRelTop = uposCursor.Y
40   mlngWidth = UserControl.Width
50   mlngHeight = UserControl.Height
End Sub

' Yorgi's 4Matz [Feb 28,2007 23:58:50] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

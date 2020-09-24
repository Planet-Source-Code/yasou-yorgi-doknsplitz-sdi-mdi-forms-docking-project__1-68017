VERSION 5.00
Begin VB.UserControl ctlButton
ClientHeight    =   255
ClientLeft      =   0
ClientTop       =   0
ClientWidth     =   255
MaskColor       =   &H00FFFFFF&
ScaleHeight     =   17
ScaleMode       =   3                             'Pixel
ScaleWidth      =   17
Begin VB.Image imgButton
Height          =   225
Left            =   30
Picture         =   "ctlButton.ctx":0000
Stretch         =   -1                            'True
Top             =   30
Width           =   225
End
End
Attribute VB_Name = "ctlButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** File Name   : ctlButton.ctl                                               **
'** Description : A custom button ActiveX control with hover effect           **
'** Usage       : As close button in ctlTitleBar ActiveX control              **
'** Dependencies: mdlAPI                                                      **
'** Members     :                                                             **
'**   * Collections: -                                                        **
'**   * Objects    : -                                                        **
'**   * Property   : Picture (def. r/w)                                       **
'**   * Methods    : -                                                        **
'**   * Events     : Click                                                    **
'** Last modified on November 13, 2003                                        **
'*******************************************************************************
Option Explicit
'--- Property Variables
Private mspPicture                     As StdPicture
'--- PropBag Names
Private Const mconPicture              As String = "Picture"
'-------------------------------
' ActiveX Control Custom Events
'-------------------------------
'Description- Occurs when the user presses and then releases a mouse button over
'             the control
Public Event Click()
'--------------------------------------------
' ActiveX Control Constructor and Destructor
'--------------------------------------------
'*****************************************************
Private Sub imgButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
   ' Purpose    - Refer to imgCloseBtn_MouseDown
   ' Inputs     - Button, Shift, X, Y
10   UserControl_MouseDown Button, Shift, X, Y
End Sub
'*****************************************************
Private Sub imgButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
   ' Purpose    - Refer to imgCloseBtn_MouseDown
   ' Inputs     - Button, Shift, X, Y
10   UserControl_MouseMove Button, Shift, X, Y
End Sub
'*****************************************************
Private Sub imgButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
   ' Purpose    - Refer to imgCloseBtn_MouseDown
   ' Inputs     - Button, Shift, X, Y
10   UserControl_MouseUp Button, Shift, X, Y
End Sub
'*****************************************************
Public Property Set Picture(spPicture As StdPicture)
'*****************************************************
   ' Purpose    - Sets a graphic to be displayed in the imgButton
   ' Input      - spPicture (the new Picture property value)
10   Set mspPicture = spPicture
20   Set imgButton.Picture = mspPicture
30   PropertyChanged mconPicture
End Property
'*****************************************************
Public Property Get Picture() As StdPicture
Attribute Picture.VB_UserMemId = 0
'*****************************************************
   ' Purpose    - Returns a graphic to be displayed in the imgButton
10   Set Picture = mspPicture
End Property
'*****************************************************
Private Sub UserControl_Initialize()
'*****************************************************
10   Set mspPicture = New StdPicture
End Sub
'*****************************************************
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
   ' Purpose    - Draws pressed effect on the control
   ' Inputs     - Button, Shift, X, Y
   Dim rec                 As RECT               'rectangle area at where to draw the pressed effect
10   If Button = vbLeftButton Then
20      mdlAPI.SetRect rec, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
30      DrawEdge UserControl.hdc, rec, mdlAPI.EDGE_SUNKEN, mdlAPI.BF_RECT
40      End If
End Sub
'*****************************************************
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
   ' Purpose    - Draws hover effect on the control
   ' Inputs     - Button, Shift, X, Y
   '  whether the mouse pointer is on the control
   Dim pos                 As POINTAPI           'used with mdlAPI.WindowFromPoint to determine
   Dim rec                 As RECT               'rectangle area at where to draw the hover effect
10   If Button = 0 Then
20      If mdlAPI.GetCapture() <> UserControl.hWnd Then
30         mdlAPI.SetCapture UserControl.hWnd
40         mdlAPI.SetRect rec, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
50         DrawEdge UserControl.hdc, rec, mdlAPI.EDGE_RAISED, mdlAPI.BF_RECT
60      Else
70         pos.X = X
80         pos.Y = Y
90         mdlAPI.ClientToScreen UserControl.hWnd, pos
100         If mdlAPI.WindowFromPoint(pos.X, pos.Y) <> UserControl.hWnd Then
110            UserControl.Cls
120            mdlAPI.ReleaseCapture
130            End If
140         End If
150      End If
End Sub
'*****************************************************
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
   ' Purposes   - * Clear all effect on the control
   '              * Raise click event if the mouse pointer is on the control
   ' Inputs     - Button, Shift, X, Y
   Dim pos                 As POINTAPI
10   If Button = vbLeftButton Then UserControl.Cls
20   If (Button = vbLeftButton) And (0 <= X) And (X <= UserControl.ScaleWidth) And (0 <= Y) And (Y <= UserControl.ScaleHeight) Then RaiseEvent Click
End Sub
'*****************************************************
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'*****************************************************
10   Set mspPicture = PropBag.ReadProperty(Name:=mconPicture, DefaultValue:=Nothing)
End Sub
'*****************************************************
Private Sub UserControl_Resize()
'*****************************************************
   ' Purpose    - Adjusts the imgButton size to match the control size
10   imgButton.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub
'*****************************************************
Private Sub UserControl_Terminate()
'*****************************************************
10   Set mspPicture = Nothing
End Sub
'*****************************************************
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'*****************************************************
10   PropBag.WriteProperty Name:=mconPicture, Value:=mspPicture, DefaultValue:=Nothing
End Sub

' Yorgi's 4Matz [Feb 28,2007 23:58:50] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

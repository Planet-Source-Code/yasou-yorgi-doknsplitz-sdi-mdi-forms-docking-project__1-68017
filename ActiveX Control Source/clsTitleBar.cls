VERSION 1.0 CLASS
BEGIN
MultiUse = -1                                     'True
Persistable = 0                                   'NotPersistable
DataBindingBehavior = 0                           'vbNone
DataSourceBehavior  = 0                           'vbNone
MTSTransactionMode  = 0                           'NotAnMTSObject
END
Attribute VB_Name = "TitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
'** File Name     - TitleBar.cls                                              **
'** Description   - An abstract class to provide the template for the         **
'**                 TitleBar interface                                        **
'*******************************************************************************
'** 01/11/07 v1.1.43 Yorgi- Add TitleBar properties:TBarType,Position
Option Explicit
'--- Class Property
Public CloseVisible                    As Boolean
Public Height                          As Long
Public Visible                         As Boolean
Public TBarType                        As TBarTypes
Public Position                        As TBarOrientation

' Yorgi's 4Matz [Feb 28,2007 23:58:51] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

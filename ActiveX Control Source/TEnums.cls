VERSION 1.0 CLASS
BEGIN
MultiUse = -1                                     'True
Persistable = 0                                   'NotPersistable
DataBindingBehavior = 0                           'vbNone
DataSourceBehavior  = 0                           'vbNone
MTSTransactionMode  = 0                           'NotAnMTSObject
END
Attribute VB_Name = "TEnums"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'TBarType constants
Public Enum TBarTypes
   TBT_SINGLESTRIPE = &H1
   TBT_GRADIENT = &H2
   TBT_DOUBLESTRIPE = &H4
   TBT_DEFAULT = TBT_DOUBLESTRIPE Or TBT_GRADIENT
End Enum
'TBarType constants
Public Enum TBarOrientation
   TBO_HORIZONTAL = &H0
   TBO_VERTICAL = &H1
End Enum
Public Enum eDockState
   DS_NotReady = 0
   DS_Docked = 1
   DS_UnDocked
End Enum
' align properties for each control
Public Enum eDAlignProperty
   DAlignNone = 0                                 ' Floating host not implemented
   DAlignTop = 1                                  ' Top host
   DAlignBottom = 2                               ' Bottom host
   DAlignLeft = 3                                 ' Left Host
   DAlignRight = 4                                ' Right Host
End Enum
' styles for the docked form
Public Enum eDockStyles
   DSLeft = &H2                                   ' Allow dock left
   DSRight = &H4                                  ' Allow dock right
   DSBottom = &H8                                 ' Allow dock bottom
   DSTop = &H10                                   ' Allow dock top
   DSFloat = &H20                                 ' Allow floating
End Enum
'-- end code

' Yorgi's 4Matz [Feb 28,2007 23:58:51] sort=subs,vars;renum=procs,10;comments=50,50;AsType=40,25

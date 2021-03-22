VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EMsgResponse
    emrConsume = 0       ' Process instead of original WindowProc
    emrPreprocess = 1    ' Process before original WindowProc
    emrPostProcess = 2  ' Process after original WindowProc
End Enum

Public Function MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    '
End Function

Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
Attribute WindowProc.VB_Description = "Raised whenever a message you have subclassed is sent."
    '
End Function



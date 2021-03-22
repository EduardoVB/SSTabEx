VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSubclass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub AttachMessage(iwp As ISubclass, ByVal hWnd As Long, ByVal iMsg As Long)
Attribute AttachMessage.VB_Description = "Start subclassing for a particular message."
    mSubclass.AttachMessage iwp, hWnd, iMsg
End Sub

Public Sub DetachMessage(iwp As ISubclass, ByVal hWnd As Long, ByVal iMsg As Long)
Attribute DetachMessage.VB_Description = "Stop subclassing for a particular message."
    mSubclass.DetachMessage iwp, hWnd, iMsg
End Sub

Public Property Get CurrentMessage() As Long
Attribute CurrentMessage.VB_Description = "Returns the message currently being worked on.  Only valid in the MsgResponse and WindowProc items of your implemented interface."
   CurrentMessage = mSubclass.CurrentMessage
End Property

Public Function CallOldWindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Attribute CallOldWindowProc.VB_Description = "Calls the window procedure for this handle which was installed before you added the subclass."
   CallOldWindowProc = mSubclass.CallOldWindowProc(hWnd, iMsg, wParam, lParam)
End Function

Public Function MyGetProp(nHwnd As Long, nPropertyName As String) As Long
    MyGetProp = mPropsDB.MyGetProp(nHwnd, nPropertyName)
End Function

Public Function MyRemoveProp(nHwnd As Long, nPropertyName As String) As Long
    MyRemoveProp = mPropsDB.MyRemoveProp(nHwnd, nPropertyName)
End Function

Public Function MySetProp(nHwnd As Long, nPropertyName As String, nData As Long) As Long
    MySetProp = mPropsDB.MySetProp(nHwnd, nPropertyName, nData)
End Function

Public Function GetMessageName(nMsg As Long) As String
    GetMessageName = mSubclass.GetMessageName(nMsg)
End Function

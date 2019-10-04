VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pocus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Function BringWindowToTop(hwnd As Long) As Long
    BringWindowToTop = Focus.BringWindowToTop(hwnd)
End Function

Function FindWindow(lpClassName As String, lpWindowName As String) As Long
    FindWindow = Focus.FindWindow(StrPtr(lpClassName), StrPtr(lpWindowName))
End Function

Function FindWindowVariant(lpClassName As Variant, lpWindowName As Variant) As Long
    FindWindowVariant = Focus.FindWindow(StrPtr(CStr(lpClassName)), StrPtr(CStr(lpWindowName)))
End Function

Function SetForegroundWindow(hwnd As Long) As Long
    SetForegroundWindow = Focus.SetForegroundWindow(hwnd)
End Function


Attribute VB_Name = "Focus"
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


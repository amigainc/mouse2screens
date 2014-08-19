Attribute VB_Name = "Module1"
Public Type POINTAPI
        x As Long
        y As Long
End Type
Public m_CursorPos As POINTAPI

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

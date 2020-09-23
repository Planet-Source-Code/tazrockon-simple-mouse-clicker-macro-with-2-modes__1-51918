Attribute VB_Name = "Module1"
Declare Function GetCursorPos& Lib "user32" (lpPoint As PointAPI)
Type PointAPI
     x As Long
     y As Long
End Type
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4

Sub MouseMove(xP As Long, yP As Long)
Dim move
move = SetCursorPos(xP, yP)
End Sub

Sub LeftClick(xP As Long, yP As Long)
mouse_event MOUSEEVENTF_LEFTDOWN, xP, yP, 0, 0
mouse_event MOUSEEVENTF_LEFTUP, xP, yP, 0, 0
End Sub

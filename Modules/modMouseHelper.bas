Attribute VB_Name = "MouseHelper"
Option Explicit

Public Declare Function MouseHelper_GetAsyncKeyState _
               Lib "Coredll" _
               Alias "GetAsyncKeyState" (ByVal vKey As Long) As Integer

Public Declare Function MouseHelper_GetSystemMetrics _
               Lib "Coredll" _
               Alias "GetSystemMetrics" (ByVal nIndex As Long) As Integer

Public Const VK_LBUTTON     As Long = &H1

Public Const VK_RBUTTON     As Long = &H2

Private Const SM_SWAPBUTTON As Long = 23

Public Function MouseHelper_IsRightMouseButtonDown() As Boolean

    Dim lngButton As Long

    If MouseHelper_GetSystemMetrics(SM_SWAPBUTTON) <> 0 Then
        lngButton = VK_LBUTTON
    Else
        lngButton = VK_RBUTTON
    End If

    MouseHelper_IsRightMouseButtonDown = MouseHelper_GetAsyncKeyState(lngButton) <> 0

End Function

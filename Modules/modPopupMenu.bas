Attribute VB_Name = "PopupMenu"
Option Explicit

Public Declare Function PopupMenu_CommandBar_GetMenu _
               Lib "Commctrl" _
               Alias "CommandBar_GetMenu" (ByVal hwndCB As Long, _
                                           ByVal iButton As Long) As Long

Public Declare Function PopupMenu_GetSubMenu _
               Lib "Coredll" _
               Alias "GetSubMenu" (ByVal hmenu As Long, _
                                   ByVal nPos As Long) As Long

Public Declare Function PopupMenu_GetWindow _
               Lib "Coredll" _
               Alias "GetWindow" (ByVal hWnd As Long, _
                                  ByVal wCmd As Long) As Long

Public Declare Function PopupMenu_GetClassName _
               Lib "Coredll" _
               Alias "GetClassNameW" (ByVal hWnd As Long, _
                                      ByVal lpClassName As String, _
                                      ByVal nMaxCount As Long) As Long

Public Declare Function PopupMenu_TrackPopupMenuEx _
               Lib "Coredll" _
               Alias "TrackPopupMenuEx" (ByVal hmenu As Long, _
                                         ByVal uFlags As Long, _
                                         ByVal x As Long, _
                                         ByVal y As Long, _
                                         ByVal hWnd As Long, _
                                         ByVal lptpm As Long) As Long

Public Declare Function PopupMenu_GetMessagePos _
               Lib "Coredll" _
               Alias "GetMessagePos" () As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GetWindow constants.

Private Const GW_CHILD        As Long = 5

Private Const GW_HWNDFIRST    As Long = 0

Private Const GW_HWNDLAST     As Long = 1

Private Const GW_HWNDNEXT     As Long = 2

Private Const GW_HWNDPREV     As Long = 3

Private Const GW_OWNER        As Long = 4

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Use one of the following flags to specify how the function positions the shortcut menu horizontally.

Public Const TPM_CENTERALIGN  As Long = &H4

Public Const TPM_LEFTALIGN    As Long = &H0

Public Const TPM_RIGHTALIGN   As Long = &H8

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Use one of the following flags to specify how the function positions the shortcut menu vertically.

Public Const TPM_BOTTOMALIGN  As Long = &H20

Public Const TPM_TOPALIGN     As Long = &H0

Public Const TPM_VCENTERALIGN As Long = &H10

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Use the following flags to control discovery of the user selection without having to set up a parent window for the menu.

Public Const TPM_NONOTIFY     As Long = &H80

Public Const TPM_RETURNCMD    As Long = &H100

Public Function PopupMenu_Show(ByRef Form As Form, _
                               ByVal MenuIndex As Long, _
                               ByVal ItemIndex As Long, _
                               ByVal Flags As Long) As Long
    
    Dim lngPos As Long

    lngPos = PopupMenu_GetMessagePos()
    
    Dim x As Long, y        As Long
    
    x = lngPos And &HFFF
    y = (lngPos And &HFF00) \ &H10000
    
    PopupMenu_Show = PopupMenu_ShowAt(Form, MenuIndex, ItemIndex, x, y, Flags)
    
End Function

Public Function PopupMenu_ShowAt(ByRef Form As Form, _
                                 ByVal MenuIndex As Long, _
                                 ByVal ItemIndex As Long, _
                                 ByVal x As Long, _
                                 ByVal y As Long, _
                                 ByVal Flags As Long) As Long

    Dim lngMessageWindow As Long

    lngMessageWindow = PopupMenu_GetCommandBarMessageWindowHandle(Form)

    Dim lngCommandBar As Long

    lngCommandBar = PopupMenu_GetCommandBarHandle(lngMessageWindow)
    
    Dim lngMenu As Long

    lngMenu = PopupMenu_CommandBar_GetMenu(lngCommandBar, MenuIndex)

    Dim lngSubMenu As Long

    lngSubMenu = PopupMenu_GetSubMenu(lngMenu, ItemIndex)
    
    PopupMenu_ShowAt = PopupMenu_TrackPopupMenuEx(lngSubMenu, Flags, x, y, lngMessageWindow, 0)

End Function

Private Function PopupMenu_GetCommandBarMessageWindowHandle(ByRef Form As Form) As Long

    Dim lngChild As Long

    lngChild = PopupMenu_GetWindow(Form.hWnd, GW_CHILD)

    While lngChild <> 0

        Dim lngSubChild As Long

        lngSubChild = PopupMenu_GetWindow(lngChild, GW_CHILD)

        If lngSubChild <> 0 Then
        
            Dim strClassName As String

            strClassName = String(Len("ToolbarWindow32"), vbNullChar)
            
            PopupMenu_GetClassName lngSubChild, strClassName, Len(strClassName) + 1

            If strClassName = "ToolbarWindow32" Then
                PopupMenu_GetCommandBarMessageWindowHandle = lngChild

                Exit Function

            End If

        End If

        lngChild = PopupMenu_GetWindow(lngChild, GW_HWNDNEXT)

    Wend

End Function

Private Function PopupMenu_GetCommandBarHandle(ByVal HwndMessage As Long) As Long
    PopupMenu_GetCommandBarHandle = PopupMenu_GetWindow(HwndMessage, GW_CHILD)
End Function




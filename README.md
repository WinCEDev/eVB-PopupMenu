# eVB-PopupMenu

This module enables you to display a PopupMenu from your eVB application.

## Usage

Simply add the *PopupMenu.bas* module and a CommandBar to your project. Set up the CommandBar as you normally would by adding the menu items you need.

When you want to display any of these items as a popup menu, simply call `PopupMenu_Show` or `PopupMenu_ShowAt`, specifying the index of the desired menu item.

A complete application could look like this:

```vb
Private Const MENU_INDEX As Long = 0 'The first menu in the CommandBar.
Private Const MENUITEM_INDEX As Long = 0 'The 'File' menu item.

Private Sub Form_Load()

    Dim objMenuBar As CommandBarMenuBar
    Set objMenuBar = CommandBar.Controls.Add(cbrMenuBar, "Menu")
    Dim objMenuItem As CommandbarLib.Item, objSubMenuItem As CommandbarLib.Item
    
    'File Menu
    Set objMenuItem = objMenuBar.Items.Add(, "File", "&File")
    objMenuItem.SubItems.Add , "Exit", "&Exit"

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Shift = 4 Then 'Detect if Alt is pressed.
        PopupMenu_Show Me, MENU_INDEX, MENUITEM_INDEX, 0 'Displays the 'File' menu.
    End If

End Sub
```

This will display the contents of the *File* menu whenever the user taps the form while holding the <kbd>Alt</kbd> key. By default, the menu is displayed at the location of the tap. If you want control over where to display the PopupMenu, you can use `PopupMenu_ShowAt`.

### Function Parameters

#### PopupMenu_Show

The `PopupMenu_Show` function takes the following parameters:

1. **Form** *(Form)*
  This should be the form that owns the CommandBar containing the menu you would like to display. It does not have to be the same form that you are calling `PopupMenu_Show` from.
2. **MenuIndex** *(Long)*
  This holds the index of the MenuBar item inside the CommandBar. Since most applications only have a single menu bar, this value is usually 0.
3. **ItemIndex** *(Long)*
  This holds the index of the menu item inside the MenuBar that you want to display. For example, if you have a MenuBar with File, Edit and View items in that order, you'd specify 1 to display the Edit menu.
4. **Flags** *(Long)*
  One or more of the [flags](https://learn.microsoft.com/en-us/previous-versions/windows/embedded/aa453773(v=msdn.10)#parameters) accepted by `TrackPopupMenuEx`. The flags are defined as constants in the module, so you can reference them by name.

#### PopupMenu_ShowAt

In addition to the parameters accepted by `PopupMenu_Show`, the `PopupMenu_ShowAt` function takes two addition parameters specifying the coordinates at which to display the popup menu.

5. **x** *(Long)*
  The horizontal location of the shortcut menu, in screen coordinates.
6. **y** *(Long)*
  The vertical location of the shortcut menu, in screen coordinates.

## Remarks

It seems like the eVB mouse events do not take into account right-clicks (the `Button` parameter is always 1). If you want to support Windows CE devices that use a mouse, you can add the included *MouseHelper.bas* module. You can then modify `Form_MouseUp` as follows:

```vb
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If MouseHelper_IsRightMouseButtonDown Or Shift = 4 Then
        PopupMenu_Show Me, MENU_INDEX, MENUITEM_INDEX, 0 'Displays the 'File' menu.
    End If
End Sub
```

The example project makes use of this modification.

## Screenshots

![Screenshot showing the example application with an expanded popup version of the Edit menu.](https://github.com/WinCEDev/eVB-PopupMenu/blob/main/Screenshots/CAPT0000.png?raw=1)

![Screenshot showing the example application displaying a message box acknowledging that the user clicked the Copy item.](https://github.com/WinCEDev/eVB-PopupMenu/blob/main/Screenshots/CAPT0001.png?raw=1)

## Links

- [HPC:Factor Forum Thread](https://www.hpcfactor.com/forums/forums/thread-view.asp?tid=21013)
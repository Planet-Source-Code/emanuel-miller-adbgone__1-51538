Attribute VB_Name = "mod_Menu"
Option Explicit

Private Declare Function GetMenu Lib "user32" _
   (ByVal hwnd As Long) As Long

Private Declare Function GetSubMenu Lib "user32" _
   (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function SetMenuItemBitmaps Lib "user32" _
   (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, _
    ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

  Const MF_END = &H80


Public Sub timeout(HowLong)
    
    '' A timeout function that is rarely used, but usefull
    Dim nofreeze%
    Dim TheBeginning
    TheBeginning = Timer
    Do
        If Timer - TheBeginning >= HowLong Then Exit Sub
        nofreeze% = DoEvents()
    Loop

End Sub
Public Function add_menu_icons(frm As Form)

    Dim MainMenu As Long, FileMenu As Long
    Dim x As Long, y As Long
    
    '' Grab the menu's handle
    MainMenu = GetMenu(frm.hwnd)

    FileMenu = GetSubMenu(MainMenu, 0)
    
    For x = 0 To 9
        y = GetMenuItemID(FileMenu, x)
        ModifyMenu FileMenu, x, MF_BYPOSITION Or MF_OWNERDRAW, y, x
    Next x


End Function

'' The old way
'Public Function add_menu_icons(frm As Form)

    'Dim mHandle As Long, lRet As Long, sHandle As Long, sHandle2 As Long
    'mHandle = GetMenu(frm.Hwnd)
    'sHandle = GetSubMenu(mHandle, 0)
   '
   ' lRet = SetMenuItemBitmaps(sHandle, 0, MF_BYPOSITION, frm.ImageList1.ListImages(6).Picture, frm.ImageList1.ListImages(6).Picture)
   ' lRet = SetMenuItemBitmaps(sHandle, 2, MF_BYPOSITION, frm.ImageList1.ListImages(7).Picture, frm.ImageList1.ListImages(7).Picture)
    'lRet = SetMenuItemBitmaps(sHandle, 4, MF_BYPOSITION, frm.ImageList1.ListImages(2).Picture, frm.ImageList1.ListImages(2).Picture)
    'lRet = SetMenuItemBitmaps(sHandle, 5, MF_BYPOSITION, frm.ImageList1.ListImages(4).Picture, frm.ImageList1.ListImages(4).Picture)
    ''lRet = SetMenuItemBitmaps(sHandle, 7, MF_BYPOSITION, frm.ImageList1.ListImages(3).Picture, frm.ImageList1.ListImages(3).Picture)
    'lRet = SetMenuItemBitmaps(sHandle, 10, MF_BYPOSITION, frm.ImageList1.ListImages(5).Picture, frm.ImageList1.ListImages(5).Picture)
    'lRet = SetMenuItemBitmaps(sHandle, 12, MF_BYPOSITION, frm.ImageList1.ListImages(1).Picture, frm.ImageList1.ListImages(1).Picture)


'End Function


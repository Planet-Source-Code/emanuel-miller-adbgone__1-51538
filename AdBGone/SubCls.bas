Attribute VB_Name = "SubCls"
Public Type MEASUREITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemWidth As Long
        itemHeight As Long
        itemData As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type DRAWITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemAction As Long
        itemState As Long
        hwndItem As Long
        hDC As Long
        rcItem As RECT
        itemData As Long
End Type

Public Const ODS_SELECTED = &H1
Public Const ODS_CHECKED = &H8
Public Const SRCCOPY = &HCC0020

Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_HILITE = &H80&
Public Const MF_MENUBREAK = &H40&
Public Const MF_MOUSESELECT = &H8000&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_SEPARATOR = &H800&
Public Const MF_UNHILITE = &H0&

Public Const WM_MEASUREITEM = &H2C
Public Const WM_DRAWITEM = &H2B

Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)


Public Const GWL_WNDPROC = -4

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private oldWndProc As Long
Private WH As Long

Public Function MyWndProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

If wMsg = WM_MEASUREITEM Then
  Dim MIS As MEASUREITEMSTRUCT
      '' Get MEASUREITEMSTRUCT data into MIS
  CopyMemory MIS, ByVal lParam, Len(MIS)
            '' ItemData contains the number of the menu item

            If MIS.itemData = 0 Then
                '' Fill data
                MIS.itemHeight = frmMenu.Picture1.ScaleHeight
                MIS.itemWidth = frmMenu.Picture1.ScaleWidth - 12
            ElseIf MIS.itemData = 1 Then
                MIS.itemHeight = frmMenu.Picture2.ScaleHeight
                MIS.itemWidth = frmMenu.Picture2.ScaleWidth - 12
            ElseIf MIS.itemData = 2 Then
                MIS.itemHeight = frmMenu.Picture3.ScaleHeight
                MIS.itemWidth = frmMenu.Picture3.ScaleWidth - 12
            ElseIf MIS.itemData = 3 Then
                MIS.itemHeight = frmMenu.Picture4.ScaleHeight
                MIS.itemWidth = frmMenu.Picture4.ScaleWidth - 12
            ElseIf MIS.itemData = 4 Then
                MIS.itemHeight = frmMenu.Picture5.ScaleHeight
                MIS.itemWidth = frmMenu.Picture5.ScaleWidth - 12
            ElseIf MIS.itemData = 5 Then
                MIS.itemHeight = frmMenu.Picture6.ScaleHeight
                MIS.itemWidth = frmMenu.Picture6.ScaleWidth - 12
            ElseIf MIS.itemData = 6 Then
                MIS.itemHeight = frmMenu.Picture7.ScaleHeight
                MIS.itemWidth = frmMenu.Picture7.ScaleWidth - 12
            ElseIf MIS.itemData = 7 Then
                MIS.itemHeight = frmMenu.Picture8.ScaleHeight
                MIS.itemWidth = frmMenu.Picture8.ScaleWidth - 12
            ElseIf MIS.itemData = 8 Then
                MIS.itemHeight = frmMenu.Picture9.ScaleHeight
                MIS.itemWidth = frmMenu.Picture9.ScaleWidth - 12
            ElseIf MIS.itemData = 9 Then
                MIS.itemHeight = frmMenu.Picture10.ScaleHeight
                MIS.itemWidth = frmMenu.Picture10.ScaleWidth - 12
            End If
            
    CopyMemory ByVal lParam, MIS, Len(MIS)
     '' Return true (we have processed the message)
    MyWndProc = 1

ElseIf wMsg = WM_DRAWITEM Then
    Dim DIS As DRAWITEMSTRUCT
    '' Get DRAWITEMSTRUCT data into DIS
    CopyMemory DIS, ByVal lParam, Len(DIS)
    '' itemData contains the number of the menu item
    If DIS.itemData = 0 Then
    '' If menu is selected
    If (DIS.itemState And ODS_SELECTED) Then
        '' Copy selected picture
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture11.ScaleWidth, frmMenu.Picture11.ScaleHeight, frmMenu.Picture11.hDC, 0&, 0&, SRCCOPY
        '' Display message
    Else
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture1.ScaleWidth, frmMenu.Picture1.ScaleHeight, frmMenu.Picture1.hDC, 0&, 0&, SRCCOPY
    End If

ElseIf DIS.itemData = 1 Then
    BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture2.ScaleWidth, frmMenu.Picture2.ScaleHeight, frmMenu.Picture2.hDC, 0&, 0&, SRCCOPY
ElseIf DIS.itemData = 2 Then
    
    If (DIS.itemState And ODS_SELECTED) Then
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture13.ScaleWidth, frmMenu.Picture13.ScaleHeight, frmMenu.Picture13.hDC, 0&, 0&, SRCCOPY
    Else
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture3.ScaleWidth, frmMenu.Picture3.ScaleHeight, frmMenu.Picture3.hDC, 0&, 0&, SRCCOPY
    End If
    
ElseIf DIS.itemData = 3 Then

    If (DIS.itemState And ODS_SELECTED) Then
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture14.ScaleWidth, frmMenu.Picture14.ScaleHeight, frmMenu.Picture14.hDC, 0&, 0&, SRCCOPY
    Else
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture4.ScaleWidth, frmMenu.Picture4.ScaleHeight, frmMenu.Picture4.hDC, 0&, 0&, SRCCOPY
    End If

ElseIf DIS.itemData = 4 Then

    If (DIS.itemState And ODS_SELECTED) Then
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture15.ScaleWidth, frmMenu.Picture15.ScaleHeight, frmMenu.Picture15.hDC, 0&, 0&, SRCCOPY
    Else
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture5.ScaleWidth, frmMenu.Picture5.ScaleHeight, frmMenu.Picture5.hDC, 0&, 0&, SRCCOPY
    End If

ElseIf DIS.itemData = 5 Then

    BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture6.ScaleWidth, frmMenu.Picture6.ScaleHeight, frmMenu.Picture6.hDC, 0&, 0&, SRCCOPY

ElseIf DIS.itemData = 6 Then
    
    If (DIS.itemState And ODS_SELECTED) Then
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture17.ScaleWidth, frmMenu.Picture17.ScaleHeight, frmMenu.Picture17.hDC, 0&, 0&, SRCCOPY
    Else
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture7.ScaleWidth, frmMenu.Picture7.ScaleHeight, frmMenu.Picture7.hDC, 0&, 0&, SRCCOPY
    End If
    
ElseIf DIS.itemData = 7 Then

    If (DIS.itemState And ODS_SELECTED) Then
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture18.ScaleWidth, frmMenu.Picture18.ScaleHeight, frmMenu.Picture18.hDC, 0&, 0&, SRCCOPY
    Else
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture8.ScaleWidth, frmMenu.Picture8.ScaleHeight, frmMenu.Picture8.hDC, 0&, 0&, SRCCOPY
    End If
    
ElseIf DIS.itemData = 8 Then

    If (DIS.itemState And ODS_SELECTED) Then
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture19.ScaleWidth, frmMenu.Picture19.ScaleHeight, frmMenu.Picture19.hDC, 0&, 0&, SRCCOPY
    Else
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture9.ScaleWidth, frmMenu.Picture9.ScaleHeight, frmMenu.Picture9.hDC, 0&, 0&, SRCCOPY
    End If
    
ElseIf DIS.itemData = 9 Then

    If (DIS.itemState And ODS_SELECTED) Then
        BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture20.ScaleWidth, frmMenu.Picture20.ScaleHeight, frmMenu.Picture20.hDC, 0&, 0&, SRCCOPY
    Else
      BitBlt DIS.hDC, DIS.rcItem.Left, DIS.rcItem.Top, frmMenu.Picture10.ScaleWidth, frmMenu.Picture10.ScaleHeight, frmMenu.Picture10.hDC, 0&, 0&, SRCCOPY
    End If
    
End If
    
    MyWndProc = 1
    
Else

    MyWndProc = CallWindowProc(oldWndProc, hwnd, wMsg, wParam, lParam)
    
End If

End Function


Public Sub SubClass(WndHwnd)

WH = WndHwnd
oldWndProc = SetWindowLong(WndHwnd, GWL_WNDPROC, AddressOf MyWndProc)

End Sub

Public Sub UnSubClass()

SetWindowLong WH, GWL_WNDPROC, oldWndProc

End Sub


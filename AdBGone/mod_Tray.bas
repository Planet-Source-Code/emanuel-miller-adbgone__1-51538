Attribute VB_Name = "mod_tray"
'' These constants are used for the system tray
'' icon
Public Const WM_LBUTTONDBLCLICK = &H203
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP


Public Type NOTIFYICONDATA
    cbSize As Long
    HWND As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
    End Type
    Public VBGTray As NOTIFYICONDATA

'' Used for updating the icon
Declare Function Shell_NotifyIcon Lib "shell32" Alias _
    "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'' Used to set focus on our icon menu
Public Declare Function SetForegroundWindow Lib "user32" (ByVal HWND As Long) As Long

Sub tray_update(frm As Form, title As String)
        
    '' This function will allow us to update
    '' the tooltip text of our icon
    Dim Tic As NOTIFYICONDATA, erg As Long
    Tic.cbSize = Len(Tic)
    Tic.HWND = frm.HWND
    Tic.uId = 1&
    Tic.uFlags = NIF_TIP
    Tic.ucallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = frm.Icon
    Tic.szTip = title
    erg = Shell_NotifyIcon(NIM_MODIFY, Tic)
    
End Sub

Sub tray(frm As Form, build As String)
    
    '' This function will put our icon
    '' into the tray
    Dim Tic As NOTIFYICONDATA, erg As Long
    Tic.cbSize = Len(Tic)
    Tic.HWND = frm.HWND
    Tic.uId = 1&
    Tic.uFlags = NIF_DOALL
    Tic.ucallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = frm.Icon
    Tic.szTip = build
    erg = Shell_NotifyIcon(NIM_ADD, Tic)

End Sub


Sub trayclose(frm As Form)
    
    '' This function will remove our icon
    '' from the tray
    Dim Tic As NOTIFYICONDATA, erg As Long
    Tic.cbSize = Len(Tic)
    Tic.HWND = frm.HWND
    Tic.uId = 1&
    erg = Shell_NotifyIcon(NIM_DELETE, Tic)

End Sub



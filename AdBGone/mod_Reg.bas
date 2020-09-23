Attribute VB_Name = "mod_Reg"
'' Registration functions
Global reg_registered As Boolean
Global reg_regkey As String
Global reg_regemail As String
Global reg_count As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Function CreateFile(filename)
    On Error Resume Next
    Close #1
    Open filename For Output As #1
        Print #1, "dump"
    Close #1
    
End Function
Public Function FileReal(filename) As Boolean
    On Error GoTo Error

    If Len(Dir(filename)) > 11 Then
        FileReal = True
    Else
        FileReal = False
    End If
    Exit Function
Error:
    Exit Function
    
End Function


Public Function SetTopMostWindow(tHWND As Long, Topmost As Boolean) As Long
 If Topmost = True Then ''Make the window topmost
  SetTopMostWindow = SetWindowPos(tHWND, HWND_TOPMOST, 0, 0, 0, 0, Flags)
 Else
  SetTopMostWindow = SetWindowPos(tHWND, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
  SetTopMostWindow = False
 End If
End Function

Public Function reg_unregGetCount() As String
On Error GoTo ender
    Dim windir As String
    Dim datafile As String
    
    windir = apiGetWindowsDirectory()
    windir = Left(windir, Len(windir) - 1)
    datafile = windir & "\system32\drivers\gmreadme.txt"

    If Not FileReal(datafile) Then
        CreateFile datafile
    End If
    

    
    Dim test(4) As Byte
    Close #1
    SetAttr datafile, vbNormal
    Open datafile For Binary Access Read As 1
        Get 1, LOF(1) - 4, test()
    Close 1
    reg_unregGetCount = ""
    For x = 0 To UBound(test)
    If Chr(test(x)) = "%" Then
        For y = (x + 1) To UBound(test)
            reg_unregGetCount = reg_unregGetCount & Chr(test(y))
        Next y
        Exit Function
    End If
    Next x
    
ender:
    
    reg_unregGetCount = "0"
    
End Function

Public Function reg_unregCheckCount()

    If reg_count > 300 Then
        disable_Protection
        reg_showReg
    Else
        reg_unregSaveCount reg_count, False
    End If
    
End Function
Public Function reg_unregSaveCount(use_count As Long, check As Boolean)
On Error Resume Next
    Dim new_flag As Boolean
    
    If Val(reg_unregGetCount) = "0" Then
        new_flag = True
    Else
        new_flag = False
    End If

    Dim windir As String
    Dim datafile As String
    
    windir = apiGetWindowsDirectory()
    windir = Left(windir, Len(windir) - 1)
    datafile = windir & "\system32\drivers\gmreadme.txt"
    
    If Not FileReal(datafile) Then
        CreateFile datafile
    End If
    
    Close 1
    SetAttr datafile, vbNormal

    Open datafile For Binary Access Write As 1
    
        If new_flag Then
            If LOF(1) = 0 Then
                Put 1, 1, "%" & Format(Val(use_count), "000")
            Else
                Put 1, LOF(1), "%" & Format(Val(use_count), "000")
            End If
        ElseIf LOF(1) > 6 Then
            Put 1, LOF(1) - 6, "%" & Format(Val(use_count), "000")
        ElseIf LOF(1) < 6 Then
            Put 1, LOF(1), "%" & Format(Val(use_count), "000")
        End If
        
    Close 1
    
    If check And use_count > 300 Then
        disable_Protection
        reg_showReg
    End If
        
        
End Function

Public Function reg_enterCode()

     Call SetTopMostWindow(frmReg.hwnd, False)
     
End Function
Public Function reg_buyNow()
    
    Call SetTopMostWindow(frmReg.hwnd, False)
    Call Shell("rundll32.exe url.dll,FileProtocolHandler http://www.ccorpsoft.com/main/modules.php?op=modload&name=products&file=index&product=adbgone", vbNormalFocus)
    
End Function
Public Function reg_showReg()

    frmReg.Show

    Call SetTopMostWindow(frmReg.hwnd, True)
    On Error Resume Next
    If verify_Backupfile Then
        On Error Resume Next
        Close #1
        Close #2
        Open backup_file For Input As #1
        Open windows_hosts_file For Output As #2
            While Not EOF(1)
                Line Input #1, temp_data
                Print #2, temp_data
            Wend
        Close #1
        Close #2
        
        Kill backup_file
    End If
    
End Function

Public Function reg_loadReg()

    On Error Resume Next
    With registry_entry
        '' Set the default location
        .hkey = HKEY_CURRENT_USER
        '' Set the root key
        .KeyRoot = "Software"
        '' Select our folder
        .Subkey = build_title
        
        '' Get the version of the adbgone application that created this configuration
        reg_regkey = .GetRegistryValue("Reg_Code")
        reg_regemail = .GetRegistryValue("Reg_Mail")
    End With
    
    
End Function

Public Function reg_saveReg(reg_code As String, reg_email As String) As Boolean

On Error GoTo err_ender
    With registry_entry
        '' Set the location
        .hkey = HKEY_CURRENT_USER
                
        '' Set the root
        .KeyRoot = "Software"
        '' Select our folder
        .Subkey = build_title
        
        '' Save values to the registry
        .SetRegistryValue "Reg_Code", reg_code, REG_SZ
        .SetRegistryValue "Reg_Mail", reg_email, REG_SZ
        
    End With
    reg_regkey = reg_code
    reg_regemail = reg_email
    reg_saveReg = True
    Exit Function
err_ender:
    reg_saveReg = False
End Function

Public Function reg_checkReg() As Boolean

    reg_loadReg
    If Len(reg_regkey) = 0 Or Len(reg_regemail) = 0 Then
        reg_checkReg = False
    Else
        reg_checkReg = reg_checkCode(reg_regkey, reg_regemail)
    End If
    
End Function

Public Function reg_Encode(reg_email As String)

On Error GoTo error_ender

    Dim email_array() As String
    Dim return_data As String
    Dim code_data As String
    Dim random_string As String
    
    code_data = Replace(reg_email, "@", "-")
    code_data = Replace(code_data, ".", "-")
    
    email_array = Split(code_data, "-")
    return_data = ""
    random_string = ""
    
    For x = 0 To UBound(email_array)
        
        If Len(email_array(x)) = 1 Then GoTo error_ender
        
        For y = 1 To Len(email_array(x))
            temp_char = Mid(email_array(x), y)
            temp_char = Left(temp_char, 1)
            If y = 1 Then
                temp = Asc(temp_char) * 50 / 2
            Else
                temp = Val(temp) + Val(Asc(temp_char) * 50 / 2)
            End If
        Next y
        If x = 0 Then random_string = temp
        
        If x < 2 Then
            email_array(x) = Hex(Val(temp)) & Oct(Asc(Val(temp)))
        Else
            email_array(x) = Hex(Val(temp) + Val(random_string)) & Oct(Asc(Val(temp)))
        End If
        
        'email_array(x) = Hex(Asc(email_array(x)) * 50 / 2) & Asc(email_array(x)) & Oct(Asc(email_array(x)))

        If Len(return_data) = 0 Then
            return_data = email_array(x)
        Else
            return_data = return_data & "-" & email_array(x)
        End If
    Next x
    
    Erase email_array
    code_data = ""
    reg_Encode = return_data
    return_data = ""
    Exit Function
    
error_ender:
    reg_Encode = "1"
    
    
    
End Function


Public Function reg_checkCode(reg_code As String, reg_email As String) As Boolean
    
    If InStr(reg_email, "@") = 0 Then
        reg_checkCode = False
        Exit Function
    ElseIf InStr(reg_email, ".") = 0 Then
        reg_checkCode = False
        Exit Function
    End If
    

    check_it = reg_Encode(reg_email)

    If check_it = "1" Then
        reg_checkCode = False
        Exit Function
    ElseIf check_it = reg_code Then
        reg_checkCode = True
        Exit Function
    Else
        reg_checkCode = False
        Exit Function
    End If
        
End Function

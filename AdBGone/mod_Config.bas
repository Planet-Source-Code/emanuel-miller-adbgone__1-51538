Attribute VB_Name = "mod_Config"
'' Random variables used throughout adbgone
Global Const build_title = "AdBGone"        '' The title
Global Const build_version = "1.3"          '' The current version

Global Const localhost = "127.0.0.1"        '' local computers adress
Global reg_version As String                '' Used for configuration versions

Global block_popups_config As Boolean       '' Couldnt use just one boolean for block_popups
Global block_popups As Boolean              '' Needed two so I could have one for configuration and one for the actual flag
Global popup_count As Long                  '' Used for keeping track of the number of popups blocked
Global advertisement_count As Long          '' Used for keeping track of the number of blocked advertisements
Global use_server, server_mode As Integer   '' Should we use the advertisement blocking server, and what mode should it be in
Global server_color As Long                 '' Color to replace ads with
Global server_html As String                '' HTML to replace ads with
Global load_at_boot As Boolean              '' Do we load at boot?
Global clear_ie_cache As Boolean            '' Should we clear ie's cache?
Global first_run As Boolean                 '' Is it the first time the programs has been run
Global protect As Boolean                   '' Protection enabled/disabled
Global hosts_file As String                 '' Wheres the hosts file
Global windows_hosts_file As String         '' Wheres window's hosts file
Global backup_file As String                '' Wheres the backup hosts file
Global html_file As String                  '' Wheres the custom html file
Global pound_data As String                 '' Header of the hosts file
Global adv_block_file As String             '' The image to replace ads with
Global menu_cancel_flag As Boolean          '' Icon menu cancel flag, show or dont show
Global windows_version As New clsOS         '' Class used for getting windows version
Global registry_entry As New RegistryRoutines '' Class used for registry manipulation
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Enum colConvertType
    colRGBtoVB = 0
    colRGBtoHTML = 1
    colVBtoRGB = 2
    colVBtoHTML = 3
    colHTMLtoRGB = 4
    colHTMLtoVB = 5
End Enum


Public Function apiGetWindowsDirectory() As String
    
    '' Get the windows directory
    Dim Buffer As String
    Buffer = Space$(255)
    GetWindowsDirectory Buffer, Len(Buffer)
    apiGetWindowsDirectory = Trim(Buffer)
    
End Function

Public Function apiGetSystemDirectory() As String

    '' Get the windows\system directory
    Dim Buffer As String
    Buffer = Space$(255)
    GetSystemDirectory Buffer, Len(Buffer)
    apiGetSystemDirectory = Trim(Buffer)
    
End Function

Public Function create_Config()
    '' This function will create the default adbgone registration keys
    
    '' If theres an error goto the err_ender tag
    On Error GoTo err_ender
    
    
    With registry_entry
        '' Set the location in registry
        .hkey = HKEY_CURRENT_USER
        
        '' Set the default root folder
        .KeyRoot = "Software"
        '' Set the subkey to nothing so we can create our folder
        .Subkey = ""
        
        '' Create a folder after adbgone's build_title
        .CreateKey build_title
        
        '' Set the subkey to the newly created folder
        .Subkey = build_title
        
        '' Create all the default required keys
        .CreateKey "Run"
        .CreateKey "Enabled"
        
        '' Set and create the values of required keys
        .SetRegistryValue "First_Run", "False", REG_SZ
        .SetRegistryValue "Enabled", "False", REG_SZ
        .SetRegistryValue "Use_Server", "1", REG_SZ
        .SetRegistryValue "Server_Mode", "1", REG_SZ
        .SetRegistryValue "Background", "16777215", REG_SZ
        .SetRegistryValue "Ad_Count", "1", REG_SZ
        .SetRegistryValue "Pop_Count", "1", REG_SZ
        .SetRegistryValue "Clear_IE", "True", REG_SZ
        .SetRegistryValue "Version", build_version, REG_SZ
        .SetRegistryValue "Block_Popups", "True", REG_SZ

        '' Set a few runtime variables
        use_server = 1
        server_mode = 1
        block_popups_config = True
        server_color = 16777215
        clear_ie_cache = True
        
        '' Set the root to windows run folder
        .KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Run"
        .Subkey = ""
        
        '' Create a key telling windows to load adbgone at boot time
        .SetRegistryValue build_title, App.Path & "\" & App.EXEName & ".exe", REG_SZ
        
    End With
        
    '' If we are creating config, then yes its the first run
    first_run = True
    '' Default config is set to load at boot
    load_at_boot = True
    
    '' Return with no errors
    create_Config = 0
    Exit Function
   
err_ender:
    '' Return with an error
    create_Config = 1
    
End Function

Public Function save_Config()

    '' This function is used for saving configuration to the registry
    With registry_entry
        '' Set the location
        .hkey = HKEY_CURRENT_USER
                
        '' Set the root
        .KeyRoot = "Software"
        '' Select our folder
        .Subkey = build_title
        
        '' Save values to the registry
        .SetRegistryValue "Block_Popups", block_popups_config, REG_SZ
        .SetRegistryValue "Use_Server", use_server, REG_SZ
        .SetRegistryValue "Server_Mode", server_mode, REG_SZ
        .SetRegistryValue "Background", server_color, REG_SZ
        .SetRegistryValue "Clear_IE", clear_ie_cache, REG_SZ
        
    End With
    
End Function
Public Function load_Config()
    
    '' This function is used for loading configuration
    
    '' If theres an error goto the err_ender tag
    On Error GoTo err_ender
    
    With registry_entry
        '' Set the default location
        .hkey = HKEY_CURRENT_USER
        '' Set the root key
        .KeyRoot = "Software"
        '' Select our folder
        .Subkey = build_title
        
        '' Get the version of the adbgone application that created this configuration
        reg_version = .GetRegistryValue("Version")
        '' If there is no value, raise an error
        If Not Len(reg_version) > 0 Then GoTo err_ender
        
        '' Load required values
        first_run = .GetRegistryValue("First_Run")
        protect = .GetRegistryValue("Enabled")
        use_server = .GetRegistryValue("Use_Server")
        reg_version = .GetRegistryValue("Version")
        block_popups_config = .GetRegistryValue("Block_Popups")
        clear_ie_cache = .GetRegistryValue("Clear_IE")
        server_color = Val(.GetRegistryValue("Background"))

        server_mode = .GetRegistryValue("Server_Mode")
        advertisement_count = Val(.GetRegistryValue("Ad_Count"))
        popup_count = Val(.GetRegistryValue("Pop_Count"))
        
        '' Change the location
        .KeyRoot = "Software\Microsoft\Windows\CurrentVersion"
        .Subkey = "Run"
        
        '' Check to see if were being loaded at boot time
        If Len(.GetRegistryValue(build_title)) <> 0 Then
            load_at_boot = True
        Else
            load_at_boot = False
        End If
        
        '' If we are blocking popups, set the flag to true
        If block_popups_config Then
            block_popups = True
        Else
            block_popups = False
        End If
    
    End With
    
    '' If an error is raised, ignore it
    On Error Resume Next
    '' Ensure that the file is closed
    Close #1
    '' Open the file for reading
    Open html_file For Input As #1
        '' Load the custom html
        Input #1, server_html
    Close #1
    
    '' If the file was empty, load the default value of custom html
    If Len(server_html) = 0 Then server_html = "<html></html>"
    
    '' Return with no errors
    load_Config = 0
    Exit Function
   
err_ender:
    '' Return with an error
    load_Config = 1
    
End Function

Public Function set_Windows_config()

    '' This function is used for setting the hosts file location
    win_dir = apiGetWindowsDirectory
    sys_dir = apiGetSystemDirectory
    win_dir = Left(win_dir, Len(win_dir) - 1)
    sys_dir = Left(sys_dir, Len(sys_dir) - 1)
    
    Select Case windows_version.OS_Name
        
        Case "Windows XP":
            windows_hosts_file = sys_dir & "\drivers\etc\hosts"
        
        Case "Windows 2000":
            windows_hosts_file = sys_dir & "\drivers\etc\hosts"
        
        Case "Windows ME":
            windows_hosts_file = win_dir & "\hosts"
            
        Case "Windows 98 SE":
            windows_hosts_file = win_dir & "\hosts"
            
        Case "Windows 98":
            windows_hosts_file = win_dir & "\hosts"
            
    End Select

    
End Function




Public Function convertColor(convert As colConvertType, ParamArray strColor()) As String

    '' This function is for converting color types
    '' The only one we use is colVBtoHTML, converting vb colors to HTML hex
    Select Case convert
        Case colRGBtoHTML
            convertColor = Left("00" & Hex(strColor(0)), 2)
            convertColor = Left("00" & convertColor & Hex(strColor(1)), 2)
            convertColor = Left("00" & convertColor & Hex(strColor(2)), 2)
        
        Case colRGBtoVB
            convertColor = RGB(strColor(0), strColor(1), strColor(2))
        
        Case colVBtoRGB
            convertColor = Right("000000" & Hex(strColor(0)), 6)
            r = CByte("&h" & Mid(convertColor, 5, 2))
            g = CByte("&h" & Mid(convertColor, 3, 2))
            b = CByte("&h" & Mid(convertColor, 1, 2))
            convertColor = r & "," & g & "," & b
            
        Case colVBtoHTML
            convertColor = Right("000000" & Hex(strColor(0)), 6)
            r = Mid(convertColor, 5, 2)
            g = Mid(convertColor, 3, 2)
            b = Mid(convertColor, 1, 2)
            convertColor = r & g & b
            
        Case colHTMLtoRGB
            convertColor = Right("000000" & Hex(strColor(0)), 6)
            r = CByte("&h" & Mid(convertColor, 1, 2))
            g = CByte("&h" & Mid(convertColor, 3, 2))
            b = CByte("&h" & Mid(convertColor, 5, 2))
            convertColor = r & "," & g & "," & b
            
        Case colHTMLtoVB
            convertColor = Right("000000" & Hex(strColor(0)), 6)
            r = CByte("&h" & Mid(convertColor, 1, 2))
            g = CByte("&h" & Mid(convertColor, 3, 2))
            b = CByte("&h" & Mid(convertColor, 5, 2))
            convertColor = RGB(r, g, b)
            
    End Select
End Function

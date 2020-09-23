Attribute VB_Name = "mod_Net"
'' These constants are used for network settings
Const net_adbgone_ip = "www.ccorpsoft.com"    '' The adbgone server
Const net_mail_ip = "mail.sbcglobal.net"            '' My mail server
Const net_mail_address = "3miller@sbcglobal.net"    '' My email
Const net_data_gethost = "GET /abg_data/abg_relay HTTP/1.1" '' The string thats sent to grab the relay file
Global net_mode As Integer  '' This is used for winsock interaction
Global net_flag As Boolean  '' This is the flag letting the function know when to stop and go
Global net_data As String   '' This is used to capture data from the socket

Public Function stop_adv_host()
    
    '' If theres an error, ignore
    On Error Resume Next
    
    With frmMenu
        '' Stop the advertisement replacement host
        .wskHost.Close
    End With
    
End Function

Public Function start_adv_host()
    
    '' If replace ads with is disabled, we dont start this
    If use_server = 0 Then Exit Function
    
    '' If theres an error, ignore
    On Error Resume Next
    With frmMenu.wskHost
        '' Set the port to listen on
        .LocalPort = 80
        '' Start listening for connections
        .Listen
    End With
    
End Function

Public Function send_adv_data(requestID As Long)
        
    '' If theres an error, skip it
    On Error Resume Next
    '' Select which index to use
    With frmMenu.wskAdv(requestID)
       
        '' Select which action to take, depending on the mode
        If server_mode = 2 Then
            '' Replace ads with a solid color
            .SendData "<html><body bgcolor=#" & convertColor(colVBtoHTML, server_color) & "></html>" & Chr(13) & Chr(10)
            GoTo ender
        ElseIf server_mode = 3 Then
            '' Replace ads with user's html
            .SendData server_html & Chr(13) & Chr(10)
            GoTo ender
        End If
        
        '' If the mode isnt selected above, then we send the adbgone image
        Dim bytBuf() As Byte
        Dim intN As Long
    
        Dim t As Integer
        Dim ByteNow As Long
        Dim ChunkSize As Long
        Dim CurrentFileSize As Long
        
        t = FreeFile
        ChunkSize = 4096
        ByteNow = 0
        '' Get the size of the adbgone logo
        CurrentFileSize = FileLen(adv_block_file)
        '' Open the adbgone logo file as binary for reading
        Open adv_block_file For Binary Access Read As #t
            '' Resize the byte array
            ReDim bytBuf(1 To ChunkSize) As Byte
            '' Loop until we cant read a proper sized chunk
            Do Until (CurrentFileSize - ByteNow) < ChunkSize
                    
                    DoEvents
                    '' Read into the array
                    Get #t, ByteNow + 1, bytBuf()
                    '' Set the file position
                    ByteNow = ByteNow + ChunkSize
                    
                    DoEvents
                    '' Send the data
                    .SendData bytBuf
            
            Loop
            
            Dim LastChunkSize As Long
            '' Set the last chunk size
            LastChunkSize = CurrentFileSize - ByteNow

            '' Make sure that it is not null
            If LastChunkSize > 1 Then
                '' Resize byte array
                ReDim bytBuf(1 To LastChunkSize) As Byte
                '' Read into the byte array
                Get #t, ByteNow + 1, bytBuf()
                '' Set the position in the file
                ByteNow = ByteNow + LastChunkSize
                '' Send the chunk
                .SendData bytBuf
            End If
    '' Close the file
    Close #t
    '' Clear array
    Erase bytBuf
    End With
    '' Goto the ender goto
    GoTo ender
    
ender:
    If Not reg_registered Then
        reg_count = reg_count + 1
        reg_unregCheckCount
    End If
    
    '' Update the advertisement counter
    advertisement_count = advertisement_count + 1
    '' Update the tooltip text to display the correct amount
    '' of blocked advertisements
    tray_update frmMenu, build_title & " (enabled)" & vbCrLf & "Ads blocked: " & advertisement_count & vbCrLf & "PopUps blocked: " & popup_count

    '' Save the advertisement block count to the registry

    '' Wait a little bit
    timeout 0.1
    '' Close the connection
    frmMenu.wskAdv(requestID).Close
    '' Unload the unused socket
    Unload frmMenu.wskAdv(requestID)
    
End Function

Public Sub contribute_Hosts(website, hosts As String)
        
    '' Disable the icon menu
    menu_cancel_flag = True
    
    '' Unload the contribution form
    Unload frmContribute
    
    With frmMenu.wskContrib
        '' Show the update/status dialog
        frmUpdate.Show
        '' Update the status
        frmUpdate.lblStatus.Caption = "Connecting..."
        
        '' Set the remote host to the mail server
        .RemoteHost = net_mail_ip
        '' Set the port to the default smtp port
        .RemotePort = 25
        '' Connect...
        .Connect
        
        '' Wait until we have a connection or error
        While .State <> 7 And .State <> 9
            DoEvents
        Wend
                
        '' If we arent connected, go to the err_ender goto tag
        If Not .State = 7 Then GoTo err_ender
        
        '' Update the status
        frmUpdate.lblStatus.Caption = "Connected, sending data..."
        
        '' Start sending data to the mail server
        .SendData "helo " & .LocalIP & Chr(13) & Chr(10)
        timeout 0.1
        .SendData "rset" & Chr(13) & Chr(10)
        timeout 0.1
        .SendData "mail from: " & build_title & "@ccorpsoft.com" & Chr(13) & Chr(10)
        timeout 0.1
        .SendData "rcpt to: <" & net_mail_address & ">" & Chr(13) & Chr(10)
        timeout 0.1
        .SendData "data" & Chr(13) & Chr(10)
        timeout 0.1
        .SendData "Subject: " & build_title & " (" & build_version & ") host" & Chr(13) & Chr(10)
        .SendData "-- website --" & Chr(13) & Chr(10)
        .SendData website & Chr(13) & Chr(10)
        .SendData "-- hosts --" & Chr(13) & Chr(10)
        .SendData hosts & Chr(13) & Chr(10)
        .SendData "." & Chr(13) & Chr(10) & Chr(13) & Chr(10)
        timeout 1#
        '' Disconnect
        .Close
        '' Update status
        frmUpdate.lblStatus.Caption = "Information sent!"
        '' Wait two seconds
        timeout 2#
        '' Unload the update/status dialog
        Unload frmUpdate
        '' Enable the icon menu
        menu_cancel_flag = False
        Exit Sub
    
    End With
    
err_ender:
    '' Disconnect
    frmMenu.wskContrib.Close
    '' Update status
    frmUpdate.lblStatus.Caption = "Error sending host(s)"
    '' Wait three seconds
    timeout 3#
    '' Unload the update/status dialog
    Unload frmUpdate
    '' Enable the icon menu
    menu_cancel_flag = False
    
End Sub

Public Sub update_Hosts()
    
    '' If we have an error, goto the end_err tag
    On Error GoTo end_err
    
    '' Disable icon menu
    menu_cancel_flag = True
    
    '' Show the update/status dialog
    frmUpdate.Show
    
    With frmMenu.wskMain
        
        '' Update status
        frmUpdate.lblStatus.Caption = "Locating update server..."
        
        '' Set the remove host to the adbgone server
        .RemoteHost = net_adbgone_ip
        '' Set the default HTTP port
        .RemotePort = 80
        '' Connect
        .Connect
        
        '' Wait until a connection or error is established
        While .State <> 7 And .State <> 9
            DoEvents
        Wend
        
        '' If we arent connected, theres a problem: goto the end_err tag
        If Not .State = 7 Then GoTo end_err
        
        '' Set the the flag to false, so we have to wait
        net_flag = False
        net_mode = 1
        
        '' Notify the server that we want the relay file
        .SendData net_data_gethost & Chr(13) & Chr(10) & _
            "Host: " & net_adbgone_ip & Chr(13) & Chr(10) & Chr(13) & Chr(10)
        
        '' Wait until sockets have done there thing
        While net_flag = False
            DoEvents
        Wend
        
        '' The sockets are done, disconnect
        .Close
        
        '' Wait a second
        timeout 1#
        
        '' Temporary variables
        Dim lst_host, lst_path As String
        '' Start parsing data so we can find out where
        '' the host file is located at
        'net_data = Mid(net_data, InStr(net_data, "X-Pad"))
        Dim array_data() As String
        array_data = Split(net_data, Chr(13) + Chr(10))

        lst_host = Left(array_data(UBound(array_data) - 1), InStr(array_data(UBound(array_data) - 1), "/") - 1)
        lst_path = Mid(array_data(UBound(array_data) - 1), InStr(array_data(UBound(array_data) - 1), "/"))

        If InStr(lst_path, Chr(10)) Then
            lst_path = Left(lst_path, InStr(lst_path, Chr(10)) - 1)
        End If
        
        '' Update the status
        frmUpdate.lblStatus.Caption = "Connecting to " & lst_host & "..."
                
        '' Set the remote host to the host we retrieved from the
        '' server
        .RemoteHost = lst_host
        '' Default port
        .RemotePort = 80
        '' Connect
        .Connect
        
        '' Wait for a connection
        While .State <> 7 And .State <> 9
            DoEvents
        Wend
        
        '' No connection?! Goto the end_err tag
        If Not .State = 7 Then GoTo end_err
        
        '' Clear the data recieved
        net_data = ""
        
        '' Update status
        frmUpdate.lblStatus.Caption = "Connected to " & lst_host & ", downloading hosts..."
        
        '' Set the flag to waiting
        net_flag = False
        '' Set the mode
        net_mode = 2
        
        '' Send data telling the server we want our host file
        .SendData "GET " & lst_path & " HTTP/1.1" & Chr(13) & Chr(10) & _
                "Host: " & lst_host & Chr(13) & Chr(10) & Chr(13) & Chr(10)
             
        '' Wait for sockets to finish
        While net_flag = False
            DoEvents
        Wend
        
        '' There done, disconnect
        .Close
        
        '' Wait a bit
        timeout 0.5
        
        '' Update status
        frmUpdate.lblStatus.Caption = "Saving hosts..."
        
        '' Make sure file is closed
        Close #1
        '' Open the hosts file and dump the data recieved
        Open hosts_file For Output As #1
            Print #1, Mid(net_data, InStr(net_data, "#"))
        Close #1
        
        '' Wait a second
        timeout 1#
        
        '' Temporary variable, explained below ---
        Dim sync_prot As Integer
        sync_prot = 1
        GoTo end_func
    End With
    
end_err:
    '' Disconnect
    frmMenu.wskMain.Close
    '' Update status
    frmUpdate.lblStatus.Caption = "Error updating hosts..."
    '' Pause for three seconds
    timeout 3#
    '' Goto the end_func tag
    Unload frmUpdate
    menu_cancel_flag = False
    MsgBox "There was an error updating the hosts. Please ensure that your internet connection is up and running.", vbCritical, build_title
    Exit Sub
    
end_func:
    '' Clear the network data
    net_data = ""
    '' Unload the update/status form
    Unload frmUpdate
    '' Enable the icon menu
    menu_cancel_flag = False
    
    '' nasty hack to fix a bug in popup protection
    If first_run Then
        MsgBox build_title & " needs to be restarted for certain features to take effect." & Chr(13) & Chr(10) & "This restart will only happen once, if you update your hosts file in the future, " & build_title & " will NOT require a restart.", vbInformation, build_title
        With registry_entry
            .hkey = HKEY_CURRENT_USER
            .KeyRoot = "Software"
            .Subkey = build_title
            .SetRegistryValue "Enabled", "True", REG_SZ
        End With
        end_Main
        Exit Sub
    End If
    
    '' If sync_prot = 1 then we will restart the protection
    '' server so the downloaded hosts will be used
    If sync_prot = 1 And frmMenu.menuProton.Checked = True Then
        disable_Protection
        timeout 0.1
        enable_Protection
    End If
    
End Sub

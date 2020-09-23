Attribute VB_Name = "mod_Encode"
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

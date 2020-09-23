VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryRoutines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'' Registry manipulation routines
Option Explicit
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hkey As Long, phkResult As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal numBytes As Long)

Const REG_OPTION_VOLATILE = 1
Const REG_OPTION_NON_VOLATILE = 0
Const SYNCHRONIZE = &H100000
Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Const STANDARD_RIGHTS_ALL = &H1F0000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Const KEY_EXECUTE = (KEY_READ)
Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Const ERROR_MORE_DATA = 234
Const ERROR_NO_MORE_ITEMS = &H103
Const ERROR_KEY_NOT_FOUND = &H2

Enum DataType
 REG_SZ = &H1
 REG_EXPAND_SZ = &H2
 REG_BINARY = &H3
 REG_DWORD = &H4
 REG_MULTI_SZ = &H7
End Enum

Enum hkey
 HKEY_CLASSES_ROOT = &H80000000
 HKEY_CURRENT_USER = &H80000001
 HKEY_LOCAL_MACHINE = &H80000002
 HKEY_USERS = &H80000003
 HKEY_PERFORMANCE_DATA = &H80000004
 HKEY_CURRENT_CONFIG = &H80000005
 HKEY_DYN_DATA = &H80000006
End Enum

Dim mvarhKeySet As Long
Dim mvarKeyRoot As String
Dim mvarSubKey As String
Dim Security As SECURITY_ATTRIBUTES

Public Property Get hkey() As hkey
    
    hkey = mvarhKeySet

End Property

Public Property Let hkey(ByVal vData As hkey)
    
    mvarhKeySet = vData

End Property

Public Property Get KeyRoot() As String
    
    KeyRoot = mvarKeyRoot

End Property

Public Property Let KeyRoot(ByVal vData As String)
    
    mvarKeyRoot = vData

End Property

Public Property Get Subkey() As String
    
    Subkey = mvarSubKey

End Property

Public Property Let Subkey(ByVal vData As String)
    
    mvarSubKey = vData

End Property

Public Function KeyExists() As Boolean
    
    Dim handle As Long
    Dim ret As Long
        
    If RegOpenKeyEx(mvarhKeySet, mvarKeyRoot & "\" & mvarSubKey, 0, KEY_READ, handle) Then
        KeyExists = False
        Exit Function
    End If
        
    KeyExists = True
    
End Function

Public Function CreateKey(Optional KeyName As Variant) As String
    
    Dim handle As Long
    Dim disp As Long
    Dim RetVal As Long
    
    KeyName = IIf(IsMissing(KeyName), mvarSubKey, CStr(KeyName))
    RetVal = RegCreateKeyEx(mvarhKeySet, mvarKeyRoot & "\" & KeyName, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, Security, handle, disp)
    
    If RetVal Then Exit Function
    
    RegCloseKey (handle)
    CreateKey = RetVal
    
End Function

Public Function DeleteKey(KeyName As String) As Long

    Dim RetVal As Long
    Dim handle As Long
    
    RetVal = RegDeleteKey(mvarhKeySet, mvarKeyRoot & "\" & KeyName)
    
    If RetVal Then Exit Function
    
    RegCloseKey (handle)
    DeleteKey = RetVal
    
End Function

Public Function DeleteValue(ValueName As String) As Long
    Dim RetVal As Long
    Dim handle As Long
    
    RetVal = RegOpenKeyEx(mvarhKeySet, mvarKeyRoot & "\" & mvarSubKey, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, handle)
    
    '' Failed
    If RetVal <> 0 Then
        DeleteValue = RetVal
        Exit Function
    End If
    
    DeleteValue = RegDeleteValue(handle, ValueName)
    RegCloseKey (handle)
    
End Function

Public Function EnumRegistryKeys(ByVal hkey As hkey, ByVal KeyName As String) As Collection
    
    Dim handle As Long
    Dim length As Long
    Dim Index As Long
    Dim subkeyName As String
    Dim fFiletime As FILETIME
    
    '' Initialize the result collection
    Set EnumRegistryKeys = New Collection

    '' Open the key, exit if not found
    If Len(KeyName) Then
        If RegOpenKeyEx(hkey, KeyName, 0, KEY_READ, handle) Then Exit Function
        hkey = handle
    End If
         
    Do
        '' Set the max length of a key
        length = 260
        subkeyName = Space$(length)
        '' get the N-th key, exit the loop if not found
        If RegEnumKeyEx(hkey, Index, subkeyName, length, 0, "", vbNull, fFiletime) = ERROR_NO_MORE_ITEMS Then Exit Do
             '' Add to the result collection
             subkeyName = Left$(subkeyName, InStr(subkeyName, vbNullChar) - 1)
             EnumRegistryKeys.Add subkeyName, subkeyName
             '' Set index for next key
             Index = Index + 1
         Loop
        
         '' Close the key
         If handle Then RegCloseKey handle
        
End Function

Function EnumRegistryValues(ByVal hkey As hkey, ByVal KeyName As String) As _
    Collection
    
    '' Temporary variables
    Dim handle As Long
    Dim Index As Long
    Dim valueType As Long
    Dim name As String
    Dim nameLen As Long
    Dim resLong As Long
    Dim resString As String
    Dim length As Long
    Dim valueInfo(0 To 1) As Variant
    Dim RetVal As Long
    Dim i As Integer
    Dim vTemp As Variant
    
    '' Initialize
    Set EnumRegistryValues = New Collection
    
    '' Open the key, exit if not found.
    If Len(KeyName) Then
        If RegOpenKeyEx(hkey, KeyName, 0, KEY_READ, handle) Then Exit Function
        hkey = handle
    End If
    
    Do
        '' Set max length of the key
        nameLen = 260
        name = Space$(nameLen)
        length = 4096
        ReDim resBinary(0 To length - 1) As Byte
        
        '' Read the value's name and data
        '' Exit the loop if not found
        RetVal = RegEnumValue(hkey, Index, name, nameLen, ByVal 0&, valueType, _
            resBinary(0), length)
        
        '' Resize the buffer if you need more space
        If RetVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To length - 1) As Byte
            RetVal = RegEnumValue(hkey, Index, name, nameLen, ByVal 0&, _
                valueType, resBinary(0), length)
        End If
        
        '' Exit the loop if any other error
        If RetVal Then Exit Do
        
        '' Retrieve the value's name
        valueInfo(0) = Left$(name, nameLen)
        
        '' Return a value corresponding to the value type
        Select Case valueType
            
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
            
            Case REG_SZ
                '' Copy everything but the trailing null char
                If length <> 0 Then
                    resString = Space$(length - 1)
                    CopyMemory ByVal resString, resBinary(0), length - 1
                    valueInfo(1) = resString
                Else
                    valueInfo(1) = ""
                End If
                
            Case REG_EXPAND_SZ
                '' Copy everything but the trailing null char
                '' Expand the environment variable to it's value
                '' Ignore a Blank String
                If length <> 0 Then
                    resString = Space$(length - 1)
                    CopyMemory ByVal resString, resBinary(0), length - 1
                    length = ExpandEnvironmentStrings(resString, resString, Len(resString))
                    valueInfo(1) = TrimNull(resString)
                Else
                    valueInfo(1) = ""
                End If

            Case REG_BINARY
                '' Shrink the buffer if necessary
                If length < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To length - 1) As Byte
                End If
                    '' Convert the string
                    For i = 0 To UBound(resBinary)
                         resString = resString & " " & Format(Trim(Hex(resBinary(i))), "0#")
                    Next i
                    valueInfo(1) = LTrim(resString) 'Get rid of leading space
            
            Case REG_MULTI_SZ
                '' Copy everything but the two trailing null chars
                resString = Space$(length - 2)
                CopyMemory ByVal resString, resBinary(0), length - 2
                
                resString = Replace(resString, vbNullChar, ",", , , vbBinaryCompare)
                valueInfo(1) = resString
            
            Case Else
                '' Unsupported value type, ignore
                
        End Select
        
        '' Add the array to the result collection
        EnumRegistryValues.Add valueInfo, valueInfo(0)
        
        Index = Index + 1
    Loop
   
    '' Close the key
    If handle Then RegCloseKey handle
        
End Function

Public Function GetRegistryValue(ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
    Dim handle As Long
    Dim resLong As Long
    Dim resString As String
    Dim TestString As String
    Dim resBinary() As Byte
    Dim length As Long
    Dim RetVal As Long
    Dim valueType As Long
    
        '' Prepare the default result
        GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)
        
        '' Open the key, exit if not found
        If RegOpenKeyEx(mvarhKeySet, mvarKeyRoot & "\" & mvarSubKey, REG_OPTION_NON_VOLATILE, KEY_READ, handle) Then
           '' Don 't overwrite the default value
           Exit Function
        End If
        
        length = 1024
        ReDim resBinary(0 To length - 1) As Byte
        
        '' Read the registry key
        RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
            length)
        '' if resBinary was too small, resize and try again
        If RetVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To length - 1) As Byte
            RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
                length)
        End If
        
        '' Key not found
        If RetVal = ERROR_KEY_NOT_FOUND Then
                 RegCloseKey (handle)
                 Exit Function
        End If
        
        '' Return a value
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                GetRegistryValue = resLong
            
            Case REG_SZ
                '' Copy everything but the trailing null char
                '' Ignore Blank Strings
                If length <> 0 Then
                    resString = Space$(length - 1)
                    CopyMemory ByVal resString, resBinary(0), length - 1
                    GetRegistryValue = resString
                End If
            
            Case REG_EXPAND_SZ
                '' Copy everything but the trailing null char
                '' Ignore a Blank String
                If length <> 0 Then
                    resString = Space$(length - 1)
                    CopyMemory ByVal resString, resBinary(0), length - 1
                    length = ExpandEnvironmentStrings(resString, resString, Len(resString))
                    GetRegistryValue = Left$(resString, length)
                End If
            
            Case REG_BINARY
                '' Resize the result
                If length <> UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To length - 1) As Byte
                End If
                GetRegistryValue = resBinary()
            
            Case REG_MULTI_SZ
                '' Copy everything but the 2 trailing null chars
                resString = Space$(length - 2)
                CopyMemory ByVal resString, resBinary(0), length - 2
                TestString = resString
                If Len(TrimNull(TestString)) > 0 Then GetRegistryValue = resString
                
            Case Else
                '' Unsupported value type, ignore
                
        End Select
        
        '' Close the registry key
     RegCloseKey (handle)
   
End Function

Public Function SetRegistryValue(ByVal ValueName As String, Value As Variant, DType As DataType) As Boolean

    Dim handle As Long
    Dim lngValue As Long
    Dim strValue As String
    Dim binValue() As Byte
    Dim length As Long
    Dim RetVal As Long
    
    '' Open the key, exit if not found
    If RegOpenKeyEx(hkey, mvarKeyRoot & "\" & mvarSubKey, REG_OPTION_NON_VOLATILE, KEY_WRITE, handle) Then
       SetRegistryValue = False 'CVar("Error!")
       Exit Function
    End If

    '' Select best course of action, depending on the data
    Select Case DType
    
        Case REG_DWORD
            lngValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
        Case REG_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, _
                Len(strValue))
        Case REG_BINARY
            binValue = Value
            length = UBound(binValue) - LBound(binValue) + 1
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_BINARY, _
                                   binValue(LBound(binValue)), length)
        Case REG_EXPAND_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_EXPAND_SZ, ByVal strValue, _
                Len(strValue))
        
        Case REG_MULTI_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_MULTI_SZ, ByVal strValue, _
                Len(strValue))
        
        Case Else
            '' Unsuported value
            
    End Select
    
    '' Close the key
    RegCloseKey (handle)
    
    SetRegistryValue = (RetVal = 0)
    
End Function

Public Function TrimNull(Item As String) As String

    '' Used to trim null chars
    Dim pos As Integer
    pos = InStr(Item, Chr$(0))
    
    If pos Then Item = Left$(Item, pos - 1)
    
    '' Return it
    TrimNull = Item
        
End Function

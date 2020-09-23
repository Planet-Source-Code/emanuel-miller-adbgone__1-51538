Attribute VB_Name = "mod_Cache"
Option Explicit

'' Begin API Declarations
'' structures used in the various cache functions, or conversion process
''
Private Const LMEM_FIXED As Long = &H0
Private Const LMEM_ZEROINIT As Long = &H40

Private Type FILETIME
    lLowDateTime As Long
    lHighDateTime As Long
End Type

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Type INTERNET_CACHE_ENTRY_INFO
   dwStructSize As Long
   lpszSourceUrlName As Long
   lpszLocalFileName As Long
   CacheEntryType As Long
   dwUseCount As Long
   dwHitRate As Long
   dwSizeLow As Long
   dwSizeHigh As Long
   LastModifiedTime As FILETIME
   ExpireTime As FILETIME
   LastAccessTime As FILETIME
   LastSyncTime As FILETIME
   lpHeaderInfo As Long
   dwHeaderInfoSize As Long
   lpszFileExtension As Long
   dwExemptDelta As Long
End Type

'' Cache functions
Private Declare Function FindFirstUrlCacheEntry Lib "wininet.dll" Alias "FindFirstUrlCacheEntryA" ( _
        ByVal lpszSearchPattern As String, _
        ByVal lpCacheInfo As Long, _
        lpdwFirstCacheEntryInfoBufferSize As Long) As Long
    
Private Declare Function FindNextUrlCacheEntry Lib "wininet.dll" Alias "FindNextUrlCacheEntryA" ( _
        ByVal hEnumHandle As Long, _
        ByVal lpCacheInfo As Long, _
        lpdwNextCacheEntryInfoBufferSize As Long) As Long

Private Declare Function FindCloseUrlCache Lib "wininet.dll" ( _
        ByVal hEnumHandle As Long) As Long
        
Private Declare Function GetUrlCacheEntryInfo Lib "wininet.dll" Alias "GetUrlCacheEntryInfoA" ( _
        ByVal lpszUrlName As String, _
        ByVal lpCacheInfo As Long, _
        lpdwCacheEntryInfoBufferSize As Long) As Long

Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" ( _
        ByVal lpszUrlName As String) As Long
    
Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyA" ( _
        ByVal RetVal As String, _
        ByVal Ptr As Long) As Long
        
Private Declare Function FileTimeToSystemTime Lib "kernel32" ( _
        lpFileTime As FILETIME, _
        lpSystemTime As SYSTEMTIME) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        pDest As Any, _
        pSource As Any, _
        ByVal dwLength As Long)

Private Declare Function LocalAlloc Lib "kernel32" ( _
        ByVal uFlags As Long, _
        ByVal uBytes As Long) As Long
    
Private Declare Function LocalFree Lib "kernel32" ( _
        ByVal hMem As Long) As Long

Public Declare Function lstrcpyA Lib "kernel32" ( _
        ByVal RetVal As String, ByVal Ptr As Long) As Long
                        
Public Declare Function lstrlenA Lib "kernel32" ( _
        ByVal Ptr As Any) As Long
  
  
'' Used for the find first, next functions...
Private hEnumHandle As Long

'' Holds value of CACHE_ENTRY_INFO between calls
Private ci As INTERNET_CACHE_ENTRY_INFO

'' Pointer to our data buffer
Private lPtrCI As Long

Public Function CachedEntryCacheType() As Long
    CachedEntryCacheType = ci.CacheEntryType
    
End Function


Public Function CachedEntryExpireTime() As Date
On Local Error Resume Next
    Dim dExpire As Date
    Dim stSystemTime As SYSTEMTIME
    Dim lReturnValue As Long
    
    ''Convert the filetime structure to a system time structure
    lReturnValue = FileTimeToSystemTime(ci.ExpireTime, stSystemTime)
    
    '' And THEN convert that to a visual basic date type
    With stSystemTime
        dExpire = CDate(.wMonth & "/" & .wDay & "/" & .wYear & " " & .wHour & ":" & .wMinute & ":" & .wSecond)
    End With
    
    '' And return that
    CachedEntryExpireTime = dExpire

End Function

Public Function CachedEntryFileExtension() As String
    Dim strData As String
    Dim lReturnValue As Long
    Dim iPosition As Long
    
    '' Allocate a buffer for our file extension
    '' Use bigger if necessary
    strData = Space(250)
    
    '' Now copy the data to our buffer, we have to use the function
    '' PtrToStr because of the way which the values are returned to us
    '' in the structure
    lReturnValue = PtrToStr(strData, ci.lpszFileExtension)
    
    '' If successful
    If lReturnValue Then
        '' Get the data we need (e.g. before the endline character)
        iPosition = InStr(strData, Chr(0))
        CachedEntryFileExtension = Left$(strData, iPosition - 1)
    End If
    

End Function

Public Function CachedEntryLastAccessTime() As Date
    Dim dExpire As Date
    Dim stSystemTime As SYSTEMTIME
    Dim lReturnValue As Long
    
    '' Convert the filetime structure to a system time structure
    lReturnValue = FileTimeToSystemTime(ci.LastAccessTime, stSystemTime)
    
    '' And THEN convert that to a visual basic date type
    With stSystemTime
        dExpire = CDate(.wMonth & "/" & .wDay & "/" & .wYear & " " & .wHour & ":" & .wMinute & ":" & .wSecond)
    End With
    
    '' And return that
    CachedEntryLastAccessTime = dExpire
    


End Function

Public Function CachedEntryLastModifiedTime() As Date
    Dim dExpire As Date
    Dim stSystemTime As SYSTEMTIME
    Dim lReturnValue As Long
    
    '' Convert the filetime structure to a system time structure
    lReturnValue = FileTimeToSystemTime(ci.LastModifiedTime, stSystemTime)
    
    '' And THEN convert that to a visual basic date type
    With stSystemTime
        dExpire = CDate(.wMonth & "/" & .wDay & "/" & .wYear & " " & .wHour & ":" & .wMinute & ":" & .wSecond)
    End With
    
    '' And return that
    CachedEntryLastModifiedTime = dExpire
    


End Function

Public Function CachedEntryLastSyncTime() As Date
    Dim dExpire As Date
    Dim stSystemTime As SYSTEMTIME
    Dim lReturnValue As Long
    
    '' Convert the filetime structure to a systemtime structure
    lReturnValue = FileTimeToSystemTime(ci.LastSyncTime, stSystemTime)
    
    '' And THEN convert that to a visual basic date type
    With stSystemTime
        dExpire = CDate(.wMonth & "/" & .wDay & "/" & .wYear & " " & .wHour & ":" & .wMinute & ":" & .wSecond)
    End With
    
    '' And return that
    CachedEntryLastSyncTime = dExpire
    


End Function
Public Function CachedEntryFileName() As String
    Dim strData As String
    Dim lReturnValue As Long
    Dim iPosition As Long
    
    '' Allocate a buffer for our filename
    '' Use bigger if necessary
    strData = String$(lstrlenA(ByVal ci.lpszLocalFileName), 0)
    
    '' Now copy the data to our buffer, we have to use the function
    '' PtrToStr because of the way which the values are returned to us
    '' In the structure
    lReturnValue = lstrcpyA(strData, ci.lpszLocalFileName)
    
    '' If successful then get the data we need
    If lReturnValue Then
        CachedEntryFileName = strData
    End If
    

End Function

Public Function CachedEntrySourceURL() As String
    Dim strData As String
    Dim lReturnValue As Long
    Dim iPosition As Long
    
    '' Allocate a buffer for our filename
    '' Use bigger if necessary
    strData = String$(lstrlenA(ci.lpszSourceUrlName), 0)
    
    '' Now copy the data to our buffer, we have to use the function
    '' lStrCpyA because of the way which the values are returned to us
    '' in the structure
    lReturnValue = lstrcpyA(strData, ci.lpszSourceUrlName)
    
    '' If successful then get the data we need
    If lReturnValue Then
        CachedEntrySourceURL = strData
    End If


End Function

Public Function DeleteCacheEntry(SourceUrl As String) As Boolean

    Dim lReturnValue As Long
    
    lReturnValue = DeleteUrlCacheEntry(SourceUrl)
    DeleteCacheEntry = CBool(lReturnValue)
    
End Function

Public Function FindEntryInCache(URL As String) As Boolean

    '' This function searches the cache for the cache entry corresponding to the
    '' given url, if this function returns true, call the various CachedEntry functions
    '' to return information about the given cache entry
    Dim lReturnValue As Long, lSizeOfStruct As Long
    
    '' Get the size needed for this structure
    lReturnValue = GetUrlCacheEntryInfo(URL, 0&, lSizeOfStruct)

    '' If we have memory allocated, free it
    If lPtrCI Then
        LocalFree lPtrCI
    End If

    '' lSizeOfStruct is now the size needed to allocate for this structure
    '' Allocate the memory for this structure
    lPtrCI = LocalAlloc(LMEM_FIXED, lSizeOfStruct)
    
    '' If the memory was succesfully allocated, then call the function again
    '' this time with the pointer to our memory block
    If lPtrCI Then
        '' I really don't know why we do this, but we do
        CopyMemory ByVal lPtrCI, lSizeOfStruct, 4

        '' Call the function again
        lReturnValue = GetUrlCacheEntryInfo(URL, lPtrCI, lSizeOfStruct)
        '' Copy the memory that our pointer points to into our structure
        CopyMemory ci, ByVal lPtrCI, Len(ci)
        '' And free our memory
        LocalFree lPtrCI
    End If
    '' And return the value as a boolean
    FindEntryInCache = CBool(lReturnValue)
    

    
End Function

Public Function FindFirstCacheEntry() As Boolean
    Dim lSizeOfStruct As Long
    
    '' The FindFirstUrlCacheEntry function returns a handle which can be
    '' used with subsequent calls to the FindNextUrlCacheEntry function
    
    '' First see if we have already opened a search handle, if so close it
    If hEnumHandle <> 0 Then
        FindCloseUrlCache hEnumHandle
    End If
        
    '' Call the FindFirstURL function with a 0& parameter to get the size of
    '' the structure
    hEnumHandle = FindFirstUrlCacheEntry(vbNullString, 0&, lSizeOfStruct)
    
    '' If we have memory allocated, free it
    If lPtrCI Then
        LocalFree lPtrCI
    End If

    '' lSizeOfStruct is now the size needed to allocate for this structure
    '' Allocate the memory for this structure
    lPtrCI = LocalAlloc(LMEM_FIXED, lSizeOfStruct)
    
    '' If the memory was succesfully allocated, then call the function again
    '' this time with the pointer to our memory block
    If lPtrCI Then
        '' I really don't know why we do this, but we do
        CopyMemory ByVal lPtrCI, lSizeOfStruct, 4

        '' Call the function again
        hEnumHandle = FindFirstUrlCacheEntry(ByVal vbNullString, lPtrCI, lSizeOfStruct)
        
        '' Copy the memory that our pointer points to into our structure
        CopyMemory ci, ByVal lPtrCI, Len(ci)
        
    End If
    
    '' Return whether successful
    FindFirstCacheEntry = CBool(hEnumHandle)
    

End Function

Public Function FindNextCacheEntry() As Boolean
    Dim lReturnValue As Long, lSizeOfStruct As Long
    
    If hEnumHandle <> 0 Then
        '' Obtain the size of the structure
        lReturnValue = FindNextUrlCacheEntry(hEnumHandle, 0&, lSizeOfStruct)
                
        '' If we have memory allocated, free it
        If lPtrCI Then
            LocalFree lPtrCI
        End If

        '' lSizeOfStruct is now the size needed to allocate for this structure
        '' Allocate the memory for this structure
        lPtrCI = LocalAlloc(LMEM_FIXED, lSizeOfStruct)
        
        '' If the memory was succesfully allocated, then call the function again
        '' this time with the pointer to our memory block
        If lPtrCI Then
            '' I really don't know why we do this, but we do
            CopyMemory ByVal lPtrCI, lSizeOfStruct, 4
            '' Call the function again
            lReturnValue = FindNextUrlCacheEntry(hEnumHandle, lPtrCI, lSizeOfStruct)
            '' Copy the memory that our pointer points to into our structure
            CopyMemory ci, ByVal lPtrCI, Len(ci)
        End If

        If lReturnValue <> 0 Then
            FindNextCacheEntry = CBool(lReturnValue)
        End If
        
    End If


    
End Function

Public Sub ReleaseCache()

    '' Call this before unloading if you used the cache functions
    If hEnumHandle Then
        Call FindCloseUrlCache(hEnumHandle)
    End If
    
    
End Sub



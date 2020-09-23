VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetVersionAdv Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Private Type OSVERSIONINFOEX
  dwOSVersionInfoSize As Long
  dwMajorVersion  As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
  wServicePackMajor As Integer
  wServicePackMinor As Integer
  wSuiteMask As Integer
  wProductType As Byte
  wReserved As Byte
End Type

Enum WinProdType
    VER_NT_WORKSTATION = 1
    VER_NT_DOMAIN_CONTROLLER = 2
    VER_NT_SERVER = 3
End Enum

Enum WinSuiteMask
    VER_SUITE_SMALLBUSINESS = &H1
    VER_SUITE_ENTERPRISE = &H2
    VER_SUITE_BACKOFFICE = &H4
    VER_SUITE_COMMUNICATIONS = &H8
    VER_SUITE_TERMINAL = &H10
    VER_SUITE_SMALLBUSINESS_RESTRICTED = &H20
    VER_SUITE_EMBEDDEDNT = &H40
    VER_SUITE_DATACENTER = &H80
    VER_SUITE_SINGLEUSERTS = &H100
    VER_SUITE_PERSONAL = &H200
    VER_SUITE_BLADE = &H400
End Enum

Public OS_Name As String
Public OS_Version As String
Public OS_Build As String
Public OS_Additional As String
Public OS_ServicePack
Public OS_Suite As String
Public OS_ProductType As String

Private Sub Class_Initialize()
    
    '' Setup our variables
    OS_Name = ""
    OS_Version = ""
    OS_Build = ""
    OS_Additional = ""
    OS_ServicePack = ""
    OS_Suite = ""
    OS_ProductType = ""
    Call GetOSVersionInfo
    
End Sub

Private Sub GetOSVersionInfo()

    Dim AdvOS As OSVERSIONINFOEX, NorOS As OSVERSIONINFO
    Dim Rtn As Long, AdvInfoAvail As Boolean
    
    AdvInfoAvail = False
    
    NorOS.dwOSVersionInfoSize = Len(NorOS)
    Rtn = GetVersionEx(NorOS)
    
    If Rtn <> 0 Then
        '' Success. Now check the OS version
        Select Case NorOS.dwMajorVersion
            '' Windows NT 3.x
            Case 3
                Select Case NorOS.dwPlatformId
                    '' Windows NT
                    Case 2
                        OS_Name = "Windows NT " & NorOS.dwMajorVersion & "." & NorOS.dwMinorVersion
                    Case Else
                        OS_Name = "Unknown OS Version " & NorOS.dwMajorVersion & "." & NorOS.dwMinorVersion
                End Select
            '' Windows 95/98/ME/NT4
            Case 4
                Select Case NorOS.dwMinorVersion
                    '' Windows 95/NT4
                    Case 0
                        Select Case NorOS.dwPlatformId
                            '' Windows 95
                            Case 1
                                If (InStr(1, UCase(NorOS.szCSDVersion), "C") > 0) Or _
                                   (InStr(1, UCase(NorOS.szCSDVersion), "B") > 0) Then
                                    OS_Name = "Windows 95 OSR2"
                                Else
                                    OS_Name = "Windows 95"
                                End If
                            '' Windows NT4
                            Case 2
                                OS_Name = "Windows NT 4.0"
                                '' Check if service pack 6 or higher is installed
                                If UCase(Trim(Replace(NorOS.szCSDVersion, Chr(0), ""))) >= "SERVICE PACK 6" Then
                                    AdvInfoAvail = True
                                End If
                            Case Else
                                OS_Name = "Unknown OS Version " & NorOS.dwMajorVersion & "." & NorOS.dwMinorVersion & " Platform " & NorOS.dwPlatformId
                        End Select
                    '' Windows 98
                    Case 10
                        If InStr(1, UCase(NorOS.szCSDVersion), "A") > 0 Then
                            OS_Name = "Windows 98 SE"
                        Else
                            OS_Name = "Windows 98"
                        End If
                    '' Windows ME
                    Case 90
                        OS_Name = "Windows ME"
                    Case Else
                        OS_Name = "Unknown OS Version " & NorOS.dwMajorVersion & "." & NorOS.dwMinorVersion
                End Select
            '' Windows 2000/XP/.NET Server
            Case 5
                AdvInfoAvail = True
                Select Case NorOS.dwMinorVersion
                    '' Windows 2000
                    Case 0
                        OS_Name = "Windows 2000"
                    '' Windows XP
                    Case 1
                        OS_Name = "Windows XP"
                    '' Windows .NET Server
                    Case 2
                        OS_Name = "Windows .NET"
                    Case Else
                        OS_Name = "Unknown OS Version " & NorOS.dwMajorVersion & "." & NorOS.dwMinorVersion
                End Select
            '' Unknown
            Case Else
                AdvInfoAvail = False
                OS_Name = "Unknown OS Version " & NorOS.dwMajorVersion & "." & NorOS.dwMinorVersion
        End Select
        
        '' Start building the normal info
        OS_Version = NorOS.dwMajorVersion & "." & NorOS.dwMinorVersion
        OS_Build = NorOS.dwBuildNumber
        OS_Additional = Replace(NorOS.szCSDVersion, Chr(0), "")
        
        If AdvInfoAvail Then
            
            '' Now get the advanced settings
            
            AdvOS.dwOSVersionInfoSize = Len(AdvOS)
            Rtn = GetVersionAdv(AdvOS)
            
            If Rtn <> 0 Then
                If (AdvOS.wServicePackMajor > 0) Or (AdvOS.wServicePackMinor > 0) Then
                    OS_ServicePack = "Service Pack " & AdvOS.wServicePackMajor & "." & AdvOS.wServicePackMinor
                Else
                    OS_ServicePack = ""
                End If
                OS_Suite = ""
                If AdvOS.wSuiteMask And WinSuiteMask.VER_SUITE_BACKOFFICE Then
                    OS_Suite = OS_Suite & "Microsoft BackOffice components are installed ,"
                End If
                If AdvOS.wSuiteMask And WinSuiteMask.VER_SUITE_COMMUNICATIONS Then
                    OS_Suite = OS_Suite & "Communications Server installed ,"
                End If
                If AdvOS.wSuiteMask And WinSuiteMask.VER_SUITE_DATACENTER Then
                    OS_Suite = OS_Suite & "Datacenter Server is installed ,"
                End If
                If AdvOS.wSuiteMask And WinSuiteMask.VER_SUITE_EMBEDDEDNT Then
                    OS_Suite = OS_Suite & "Embedded NT installed ,"
                End If
                If AdvOS.wSuiteMask And WinSuiteMask.VER_SUITE_SINGLEUSERTS Then
                    OS_Suite = OS_Suite & "Single User ,"
                End If
                If AdvOS.wSuiteMask And WinSuiteMask.VER_SUITE_SMALLBUSINESS Then
                    OS_Suite = OS_Suite & "Microsoft Small Business Server is installed ,"
                End If
                If AdvOS.wSuiteMask And WinSuiteMask.VER_SUITE_SMALLBUSINESS_RESTRICTED Then
                    OS_Suite = OS_Suite & "Microsoft Small Business Server is installed with the restrictive client license in force ,"
                End If
                If AdvOS.wSuiteMask And WinSuiteMask.VER_SUITE_TERMINAL Then
                    OS_Suite = OS_Suite & "Terminal Services is installed ,"
                End If
                
                If Len(OS_Suite) > 2 Then
                    OS_Suite = Left$(OS_Suite, Len(OS_Suite) - 1)
                End If
                
                Select Case AdvOS.dwMajorVersion
                    '' NT4
                    Case 4
                        Select Case AdvOS.wProductType
                            Case WinProdType.VER_NT_WORKSTATION
                                OS_ProductType = "Workstation"
                            Case WinProdType.VER_NT_DOMAIN_CONTROLLER
                                OS_ProductType = "Domain Controller"
                            Case WinProdType.VER_NT_SERVER
                                If AdvOS.wSuiteMask And WinSuiteMask.VER_SUITE_ENTERPRISE Then
                                    OS_ProductType = "Advanced Server"
                                Else
                                    OS_ProductType = "Server"
                                End If
                        End Select
                    '' 2000/XP/.NET
                    Case 5
                        Select Case AdvOS.dwMinorVersion
                            '' 2000
                            Case 0
                                Select Case AdvOS.wProductType
                                    Case WinProdType.VER_NT_WORKSTATION
                                        OS_ProductType = "Professional"
                                    Case WinProdType.VER_NT_DOMAIN_CONTROLLER
                                        OS_ProductType = "Domain Controller"
                                    Case WinProdType.VER_NT_SERVER
                                        If AdvOS.wSuiteMask And WinSuiteMask.VER_SUITE_ENTERPRISE Then
                                            OS_ProductType = "Advanced Server"
                                        Else
                                            OS_ProductType = "Server"
                                        End If
                                End Select
                            '' XP
                            Case 1
                                Select Case AdvOS.wProductType
                                    Case WinProdType.VER_NT_WORKSTATION
                                        If AdvOS.wSuiteMask And WinSuiteMask.VER_SUITE_PERSONAL Then
                                            OS_ProductType = "Home Edition"
                                        Else
                                            OS_ProductType = "Professional"
                                        End If
                                    Case WinProdType.VER_NT_DOMAIN_CONTROLLER
                                        OS_ProductType = "Domain Controller"
                                    Case WinProdType.VER_NT_SERVER
                                        OS_ProductType = "Server"
                                End Select
                            '' .NET
                            Case 2
                                Select Case AdvOS.wProductType
                                    Case WinProdType.VER_NT_WORKSTATION
                                        '' Undocumented by MS
                                        OS_ProductType = ""
                                    Case WinProdType.VER_NT_DOMAIN_CONTROLLER
                                        '' Also undocumented but must be domain controller
                                        OS_ProductType = "Domain Controller"
                                    Case WinProdType.VER_NT_SERVER
                                        If AdvOS.wSuiteMask And WinSuiteMask.VER_SUITE_DATACENTER Then
                                            OS_ProductType = "DataCenter Server"
                                        ElseIf AdvOS.wSuiteMask And WinSuiteMask.VER_SUITE_ENTERPRISE Then
                                            OS_ProductType = "Enterprise Server"
                                        ElseIf AdvOS.wSuiteMask And WinSuiteMask.VER_SUITE_BLADE Then
                                            OS_ProductType = "Web Server"
                                        Else
                                            OS_ProductType = "Server"
                                        End If
                                End Select
                            Case Else
                                '' Unknown minor version type
                                OS_ProductType = ""
                        End Select
                    Case Else
                        '' Unknown OS Type
                        OS_ProductType = ""
                End Select
            End If
        End If
    End If
End Sub


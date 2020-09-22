Attribute VB_Name = "GetOS"
Option Explicit
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Public Function GetWindowsVersion() As String
    Dim TheOS As OSVERSIONINFO
    Dim strCSDVersion As String

    TheOS.dwOSVersionInfoSize = Len(TheOS)
    GetVersionEx TheOS
    Select Case TheOS.dwPlatformId
    Case VER_PLATFORM_WIN32_WINDOWS
        If TheOS.dwMinorVersion >= 10 Then
            GetWindowsVersion = "Windows 98 version: "
        Else
            GetWindowsVersion = "Windows 95 version: "
        End If
    Case VER_PLATFORM_WIN32_NT
        GetWindowsVersion = "Windows NT version: "
    End Select
    'Extract the Additional Version Information from the string with null char terminator
    If InStr(TheOS.szCSDVersion, Chr(0)) <> 0 Then
        strCSDVersion = ": " & Left(TheOS.szCSDVersion, InStr(TheOS.szCSDVersion, Chr(0)) - 1)
    Else
        strCSDVersion = ""
    End If
    GetWindowsVersion = GetWindowsVersion & TheOS.dwMajorVersion & "." & TheOS.dwMinorVersion & " (Build " & TheOS.dwBuildNumber & strCSDVersion & ")"
End Function

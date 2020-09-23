Attribute VB_Name = "Module1"
Option Explicit
'API calls
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetComputername Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Declare Function GetDesktopWindow& Lib "user32" ()
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetTickCount& Lib "kernel32" ()
Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean


Public Const HWND_NOTOPMOST = -2 'Sets Form Notontop (More to it look below)
Public Const HWND_TOPMOST = -1 ' Sets Form Ontop (More to it look below)
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const EWX_SHUTDOWN = 1 'Part of the Option to shutdown windows
Public Const EWX_REBOOT = 2 'Part of the Option to shutdown windows
Public Const HWND_DESKTOP = 0
Public Const SW_NORMAL = 1
Public Const NIM_ADD = &H0 'Systray Crap
Public Const NIM_MODIFY = &H1 'Systray Crap
Public Const NIM_DELETE = &H2 'Systray Crap
Public Const WM_MOUSEMOVE = &H200 'Systray Crap
Public Const NIF_MESSAGE = &H1 'Systray Crap
Public Const NIF_ICON = &H2 'Systray Crap
Public Const NIF_TIP = &H4 'Systray Crap
Public Const WM_LBUTTONDBLCLK = &H203 'Systray Crap
Public Const WM_LBUTTONDOWN = &H201 'Systray Crap
Public Const WM_LBUTTONUP = &H202 'Systray Crap
Public Const WM_RBUTTONDBLCLK = &H206 'Systray Crap
Public Const WM_RBUTTONDOWN = &H204 'Systray Crap
Public Const WM_RBUTTONUP = &H205 'Systray Crap

'Menu for OSversion information
Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
'Menu for Systray information
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public SysIcon As NOTIFYICONDATA, RunningInTray As Boolean

Public Sub TimeOut(HowLong)
    'Something like a timer without a timer
    Dim TheBeginning
    Dim NoFreeze As Integer
    TheBeginning = Timer
    Do
        If Timer - TheBeginning >= HowLong Then Exit Sub
        NoFreeze% = DoEvents()
    Loop
End Sub

Public Function File_ByteConversion(NumberOfBytes As Single) As String
On Error Resume Next
    'calculate filesize with any file
    If NumberOfBytes < 1024 Then 'checks to see if its so small that it cant be converted into larger grouping
        File_ByteConversion = NumberOfBytes & " Bytes"
    End If
    
    If NumberOfBytes > 1024 Then  'Checks to see if file is big enough to convert into KB
        File_ByteConversion = Format(NumberOfBytes / 1024, "0.00") & " KB"
    End If

    If NumberOfBytes > 1024000 Then 'Checks to see if its big enough to convert into MB
        File_ByteConversion = Format(NumberOfBytes / 1024000, "###,###,##0.00") & " MB"
    End If
End Function

Public Function Ontop(FormName As Form) As Boolean
'Sets form ontop of all other programs
If Ontop = False Then
    Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
Else
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End If
End Function

Public Function getVersion() As String
'Tries to get Windows OS Version
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
With osinfo
    Select Case .dwPlatformId
        Case 1
            If .dwMinorVersion = 0 Then
                getVersion = "Windows 95"
            ElseIf .dwMinorVersion = 10 Then
                getVersion = "Windows 98"
            End If
        Case 2
            If .dwMajorVersion = 3 Then
                getVersion = "Windows NT 3.51"
            ElseIf .dwMajorVersion = 4 Then
                getVersion = "Windows NT 4.0"
            ElseIf .dwMajorVersion = 5 Then
                getVersion = "Windows XP"
            End If
        Case Else
            getVersion = "Failed"
    End Select
End With
End Function

Public Function AppDir() As String
'Sometimes this causes a problem to some people
'because they install stuff to C:\ so this makes it simple
'and error free.  Just use "APPDIR" now insted of "APP.PATH"
    If Right$(App.Path, 1) <> 1 Then
        AppDir = App.Path & "\"
    Else
        AppDir = App.Path
    End If
End Function

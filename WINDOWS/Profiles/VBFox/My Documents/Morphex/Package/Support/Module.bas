Attribute VB_Name = "Module"
Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwnd1 As Long, ByVal hwnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInstertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const FLAGS = SWP_NOSIZE Or SWP_NOMOVE
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SPI_SCREENSAVERRUNNING = 97

Public Enum CDROMState_Constants
 Closed = 1
 Opened = 2
End Enum
Public Enum CursorState_Constants
 Showing = True
 Hiding = False
End Enum
Public Enum KeyButtons_Constants
 KeysOn = False
 KeysOff = True
End Enum
Public Enum TaskBar_Constants
 IsShowing = True
 IsHiding = False
End Enum
Public Enum Desktop_Constants
 IsOn = True
 IsOff = False
End Enum
Public Enum FormTop_Constants
 IsOnTop = -1
 IsNotOnTop = -2
End Enum
Public Enum StartBar_Constants
 IsOnTaskbar = 1
 InNotOnTaskbar = 0
End Enum


Public Function Desktop(State As Desktop_Constants)
Dim DesktopHwnd As Long
Dim SetOption As Long

DesktopHwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
SetOption = IIf(State, SW_SHOW, SW_HIDE)
ShowWindow DesktopHwnd, SetOption
End Function


Public Function OpenURL(ByVal URL As String) As Long
OpenURL = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function

Public Function Cursor(State As CursorState_Constants)
Dim SendValue As Long
SendValue = ShowCursor(State)
End Function

Public Function CDRom(State As CDROMState_Constants)
Dim SendValue As Long
Select Case State
Case 1
SendValue = MciSendString("set CDAudio door closed", vbNullString, 0, 0)
Case 2
SendValue = MciSendString("set CDAudio door open", vbNullString, 0, 0)
End Select
End Function

Public Function KeyButtons(State As KeyButtons_Constants)
Dim SendValue As Long
Dim SetOption As Boolean
SendValue = SystemParametersInfo(SPI_SCREENSAVERRUNNING, State, SetOption, 0)
End Function

Public Function TaskBar(State As TaskBar_Constants)
Dim TaskbarHwnd As Long
Dim SendValue As Long
Dim SetOption As Long

SetOption = IIf(State, SW_SHOW, SW_HIDE)
TaskbarHwnd = FindWindow("Shell_TrayWnd", "")
SendValue = ShowWindow(TaskbarHwnd, SetOption)
End Function

Public Function CenterForm(FormName As Form)
FormName.Top = (Screen.Height * 0.95) / 2 - FormName.Height / 2
FormName.Left = Screen.Width / 2 - FormName.Width / 2
End Function

Public Function FormState(FormName As Form, State As FormTop_Constants)
Dim SendValue As Long
SendValue = SetWindowPos(FormName.hwnd, State, 0, 0, 0, 0, FLAGS)
End Function

Public Function StartButton(State As StartBar_Constants)
Dim SendValue As Long
Dim SetOption As Long

SetOption = FindWindow("Shell_TrayWnd", "")
SendValue = FindWindowEx(SetOption, 0, "Button", vbNullString)
ShowWindow SendValue, State
End Function

Public Function TaskbarClock(State As StartBar_Constants)
Dim SendValue As Long
Dim SetOption As Long
Dim MainOption As Long

SetOption = FindWindow("Shell_TrayWnd", "")
MainOption = FindWindowEx(SetOption, 0, "TrayNotifyWnd", vbNullString)
SendValue = FindWindowEx(MainOption, 0, "TrayClockWClass", vbNullString)
ShowWindow SendValue, State
End Function

Public Function TaskbarIcons(State As StartBar_Constants)
Dim SendValue As Long
Dim SetOption As Long

SetOption = FindWindow("Shell_TrayWnd", "")
SendValue = FindWindowEx(SetOption, 0, "TrayNotifyWnd", vbNullString)
ShowWindow SendValue, State
End Function

Public Function TaskbarPrograms(State As StartBar_Constants)
Dim SendValue As Long
Dim SetOption As Long
Dim MainOptionA As Long
Dim MainOptionB As Long

SetOption = FindWindow("Shell_TrayWnd", "")
MainOptionA = FindWindowEx(SetOption, 0, "ReBarWindow32", vbNullString)
MainOptionB = FindWindowEx(MainOptionA, 0, "MSTaskSwWClass", vbNullString)
SendValue = FindWindowEx(MainOptionB, 0, "SysTabControl32", vbNullString)
ShowWindow SendValue, State
End Function

Public Function CleanUp()
frmCleanUp.Show
frmMain.Enabled = False
CDRom Closed
Cursor Showing
Desktop IsOn
KeyButtons KeysOn
StartButton IsOnTaskbar
TaskBar IsShowing
TaskbarIcons IsOnTaskbar
TaskbarPrograms IsOnTaskbar
TaskbarClock IsOnTaskbar
End Function

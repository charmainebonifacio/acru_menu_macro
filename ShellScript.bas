Attribute VB_Name = "ShellScript"
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAcess As Long, _
    ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
'---------------------------------------------------------------------
' Date Acquired: July 4, 2013
' Source       : http://www.vbaexpress.com/forum/showthread.php?t=37457
'---------------------------------------------------------------------
' Date Edited  : July 4, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : Shell_AndWait
' Description  : This function stops the application in its tracks --
'                it doesn't repond to keyboard, mouse, etc until
'                the shelled app is finished.
' Parameters   : String, ShellWindow, Long
' Returns      : Boolean
'---------------------------------------------------------------------
Public Enum ShellTiming
    SH_IGNORE = 0 'Ignore signal
    SH_INFINITE = -1& 'Infinite timeout
    SH_PROCESS_QUERY_INFORMATION = &H400
    SH_STILL_ACTIVE = &H103
    SH_SYNCHRONIZE = &H100000
End Enum
Public Enum ShellWait
    SH_WAIT_ABANDONED = &H80&
    SH_WAIT_FAILED = -1& 'Error on call
    SH_WAIT_OBJECT_0 = 0 'Normal completion
    SH_WAIT_TIMEOUT = &H102& 'Timeout period elapsed
End Enum
Public Enum ShellWindow
    SH_HIDE = 0
    SH_SHOWNORMAL = 1 'normal with focus
    SH_SHOWMINIMIZED = 2 'minimized with focus (default in VB)
    SH_SHOWMAXIMIZED = 3 'maximized with focus
    SH_SHOWNOACTIVATE = 4 'normal without focus
    SH_SHOW = 5 'normal with focus
    SH_MINIMIZE = 6 'minimized without focus
    SH_SHOWMINNOACTIVE = 7 'minimized without focus
    SH_SHOWNA = 8 'normal without focus
    SH_RESTORE = 9 'normal with focus
End Enum
Function Shell_AndWait(ByVal CommandLine As String, _
    Optional ExecMode As ShellWindow = SH_HIDE, _
    Optional Timeout As Long = SH_INFINITE) As Boolean
    Dim ProcessID As Long
    Dim hProcess As Long
    Dim nRet As Long
    Const fdwAccess = SH_SYNCHRONIZE
    If ExecMode < SH_HIDE Or ExecMode > SH_RESTORE Then ExecMode = SH_SHOWNORMAL
    ProcessID = Shell(CommandLine, CLng(ExecMode))
    hProcess = OpenProcess(fdwAccess, False, ProcessID)
    nRet = WaitForSingleObject(hProcess, CLng(Timeout))
    Shell_AndWait = (nRet <> 0)
End Function


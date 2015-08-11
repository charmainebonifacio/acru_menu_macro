Attribute VB_Name = "WARNING"
'---------------------------------------------------------------------------------------
' Date Created : June 12, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : August 10, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : WarningMessage
' Description  : This function will notify user that tool is currently processing the
'                user request.
' Parameters   : -
' Returns      : -
'---------------------------------------------------------------------------------------
Function WarningMessage()

    Dim WarningPrompt As String
    Dim WindowTitle As String
    Dim DefaultTimer As Long
    
    DefaultTimer = 100 ' < Set Timer
    
    WindowTitle = "The Processing Zonal Statistics Tool"
    WarningPrompt = "ATTENTION." & vbCrLf
    WarningPrompt = WarningPrompt & vbCrLf
    WarningPrompt = WarningPrompt & "The macro is currently processing your request. " & vbCrLf
    WarningPrompt = WarningPrompt & vbCrLf
    WarningPrompt = WarningPrompt & "Please click [OK] to continue." & vbCrLf
    WarningPrompt = WarningPrompt & vbCrLf
    
    TimedMsgBox WarningPrompt, DefaultTimer, WindowTitle ' Call New MsgBox
    
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 13, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : August 10, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : MacroTimer
' Description  : This function will notify user how much time has elapsed to complete
'                the entire procedure.
' Parameters   : Long
' Returns      : String
'---------------------------------------------------------------------------------------
Function MacroTimer(ByVal TimeElapsed As Long) As String

    Dim NotifyUser As String
    
    NotifyUser = NotifyUser & "The program has completed. "
    NotifyUser = NotifyUser & "It took a total of " & TimeElapsed & " seconds."

    MacroTimer = NotifyUser
    
End Function

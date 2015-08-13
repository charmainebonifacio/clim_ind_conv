Attribute VB_Name = "WARNING"
'---------------------------------------------------------------------------------------
' Date Created : January 9, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : January 10, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : WarningMessage
' Description  : This function will notify user that tool is currently processing the
'                user request to split the timeseries into two new timeseries.
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
    WarningPrompt = WarningPrompt & "The macro is currently processing your request." & vbCrLf
    WarningPrompt = WarningPrompt & vbCrLf
    WarningPrompt = WarningPrompt & "Please click [OK] to continue." & vbCrLf
    WarningPrompt = WarningPrompt & vbCrLf
    
    TimedMsgBox WarningPrompt, DefaultTimer, WindowTitle ' Call New MsgBox
    
End Function
'---------------------------------------------------------------------------------------
' Date Created : January 9, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : January 10, 2015
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
    
    NotifyUser = "MACRO RUN IS SUCCESSFUL!"
    NotifyUser = NotifyUser & vbCrLf
    NotifyUser = NotifyUser & "The macro has processed grid .CSV files for mean monthly data." & vbCrLf
    NotifyUser = NotifyUser & "The macro run took a total of " & TimeElapsed & " minutes."

    MacroTimer = NotifyUser
    
End Function

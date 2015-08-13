Attribute VB_Name = "Step_1_Main"
Public objFSOlog As Object
Public logfile As TextStream
Public logtxt As String
Public appSTATUS As String
'---------------------------------------------------------------------------------------
' Date Created : August 13, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : August 13, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CLIMIND_MAIN
' Description  : This is the main function that will convert .XLSX files to .CSV
'---------------------------------------------------------------------------------------
Function CLIMIND_MAIN()

    Dim start_time As Date, end_time As Date
    Dim ProcessingTime As Long
    Dim MessageSummary As String, SummaryTitle As String
    
    Dim UserSelectedFolder As String
    Dim MAINFolder As String, TmpOUT As String, outDir As String
    Dim filesProcessed As Integer
    
    ' Initialize Variables
    SummaryTitle = "Tool Summary"
    outDir = "Output"
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
        
    '---------------------------------------------------------------------
    ' I. FIND DIRECTORY
    '---------------------------------------------------------------------
    UserSelectedFolder = GetFolder
    Debug.Print UserSelectedFolder
    If Len(UserSelectedFolder) = 0 Then GoTo Cancel
    MAINFolder = ReturnFolderName(UserSelectedFolder)
    Debug.Print MAINFolder
        
    '---------------------------------------------------------------------
    ' II. LOGFILE SETUP
    '---------------------------------------------------------------------
    TmpOUT = ReturnSubFolder(UserSelectedFolder, outDir)
    CheckOUTFolder = CheckFolderExists(TmpOUT)
    Debug.Print CheckOUTFolder
    If CheckOUTFolder = False Then NewOUT = CreateFolder(TmpOUT)

    Dim logfilename As String, logtextfile As String, logext As String
    logext = ".txt"
    logfilename = "clim_ind_conversion"
    logtextfile = SaveLogFile(TmpOUT, logfilename, logext)
    
    Set objFSOlog = CreateObject("Scripting.FileSystemObject")
    Set logfile = objFSOlog.CreateTextFile(logtextfile, True)
        
    '---------------------------------------------------------------------
    ' III. START PROGRAM
    '---------------------------------------------------------------------
    start_time = Now()
    logfile.WriteLine "[ START PROGRAM. ] "
    logfile.WriteLine ""
    logfile.WriteLine "User selected the following directory : " & UserSelectedFolder
    logfile.WriteLine ""
    logfile.WriteLine "[ PROCESSING FILE SUMMARY ]"
    filesProcessed = PROCESSFILES(UserSelectedFolder, TmpOUT)
    logfile.WriteLine " "
    If filesProcessed = 0 Then GoTo Cancel
    
    '---------------------------------------------------------------------
    ' IV. END PROGRAM
    '---------------------------------------------------------------------
    end_time = Now()
    ProcessingTime = DateDiff("s", CDate(start_time), CDate(end_time))
    MessageSummary = MacroTimer(ProcessingTime)
    MsgBox MessageSummary, vbOKOnly, SummaryTitle
    
Cancel:
    If Len(UserSelectedFolder) = 0 Then
        MsgBox "No folder selected.", vbOKOnly, SummaryTitle
    End If
    If filesProcessed = 0 Then
        end_time = Now()
        ProcessingTime = DateDiff("s", CDate(start_time), CDate(end_time))
        MessageSummary = MacroTimer(ProcessingTime)
        logfile.WriteLine "Unable to find valid files to convert."
    End If

    ' Close Log File
    logfile.WriteLine " "
    logfile.WriteLine MessageSummary
    logfile.WriteLine " "
    logfile.WriteLine "[ END PROGRAM. ] "
    logfile.Close
    Set logfile = Nothing
    Set objFSOlog = Nothing
End Function

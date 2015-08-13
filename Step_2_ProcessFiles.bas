Attribute VB_Name = "Step_2_ProcessFiles"
'---------------------------------------------------------------------
' Date Created : August 13, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : August 13, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : PROCESSFILES
' Description  : This function processes all the .XLSX files within a
'                file directory and converts the files to .CSV.
' Parameters   : String, String
' Returns      : Boolean
'---------------------------------------------------------------------
Function PROCESSFILES(ByVal fileDir As String, ByVal outDir As String) As Integer

    Dim objFolder As Object, objFSO As Object
    Dim wbSource As Workbook, SourceSheet As Worksheet
    Dim FileCounter As Long
    Dim sThisFilePath As String, sFile As String
    Dim GridName As String
    Dim VarType As String
    Dim fileExt As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Status Bar Update
    appSTATUS = "Processing file within the selected folder..."
    Application.StatusBar = appSTATUS
    logtxt = appSTATUS
    logfile.WriteLine logtxt
    
    ' Initialize Variables
    FileCounter = 0
    fileExt = ".xlsx"
    
    '-------------------------------------------------------------
    ' Check the files... which should be obvious at this point.
    '-------------------------------------------------------------
    sThisFilePath = fileDir
    If (Right(sThisFilePath, 1) <> "\") Then sThisFilePath = sThisFilePath & "\"
    sFile = Dir(sThisFilePath & "*.csv")

    '-------------------------------------------------------------
    ' Loop through all files
    '-------------------------------------------------------------
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(sThisFilePath).Files
    
    For Each objFILE In objFolder
        logtxt = objFILE
        Debug.Print objFILE
        logfile.WriteLine objFILE
        
        If UCase(Right(objFILE.Path, (Len(objFILE.Path) - InStrRev(objFILE.Path, ".")))) = UCase("xlsx") Then
            FileCounter = FileCounter + 1
            logtxt = FileCounter & " of files processed."
            Debug.Print logtxt
            logfile.WriteLine logtxt
            
            ' Open file and set it as source worksheet
            Set wbSource = Workbooks.Open(objFILE.Path)
            Set SourceSheet = wbSource.Worksheets(wbSource.Worksheets.Count)
            SourceSheet.Activate
                       
            ' Exclude file extension
            logfile.WriteLine "The workbook full filename is: " & wbSource.Name
            GridName = Replace(wbSource.Name, fileExt, "")
            logfile.WriteLine "The workbook filename is: " & GridName
            
            ' Save Changes to the Processed Files
            Call SaveAsCSV(wbSource, outDir, GridName)
            wbSource.Close SaveChanges:=False

        Else:
            logtxt = objFILE & " is not a valid file to process."
            Debug.Print logtxt
            logfile.WriteLine logtxt
        End If
    Next
    
    PROCESSFILES = FileCounter
    
Cancel:
    Set wbSource = Nothing
    Set SourceSheet = Nothing
    Set wbDest = Nothing
    Set DestSheet = Nothing
    Set objFSO = Nothing
    Set objFolder = Nothing
End Function

Attribute VB_Name = "Step_0_Start"
'---------------------------------------------------------------------
' Date Created : August 13, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : August 13, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : Start_Here
' Description  : The purpose of function is to initialize the userform.
'---------------------------------------------------------------------
Sub Start_Here()
   
    Dim myForm As UserForm1
    Dim button1 As String, button2 As String, button3 As String
    Dim button4 As String, button5 As String, button6 As String
    Dim strLabel1 As String, strLabel2 As String
    Dim strLabel3 As String, strLabel4 As String
    Dim strLabel5 As String, strLabel6 As String
    Dim strLabel7 As String, strLabel8 As String
    Dim frameLabel1 As String, frameLabel2 As String, frameLabel3 As String
    Dim userFormCaption As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    Set myForm = UserForm1
    
    ' Label Strings
    userFormCaption = "KIENZLE LAB TOOLS"
    button1 = "CONVERT FILES"
    frameLabel2 = "TOOL GUIDE"
    frameLabel3 = "HELP SECTION"
    
    strLabel1 = "THE CLIMATE INDICES CONVERSION MACRO"
    strLabel2 = "STEP 1."
    strLabel3 = "Move all .XLSX files into one directory." & vbLf
    strLabel4 = "STEP 2."
    strLabel5 = "For more information, hover mouse over button."
    
    ' UserForm Initialize
    myForm.Caption = userFormCaption
    myForm.Frame2.Caption = frameLabel2
    myForm.Frame5.Caption = frameLabel3
    myForm.Frame2.Font.Bold = True
    myForm.Frame5.Font.Bold = True
    myForm.Label1.Caption = strLabel1
    myForm.Label1.Font.Size = 21
    myForm.Label1.Font.Bold = True
    myForm.Label1.TextAlign = fmTextAlignCenter
    
    myForm.Label2 = strLabel2
    myForm.Label2.Font.Size = 13
    myForm.Label2.Font.Bold = True
    myForm.Label3 = strLabel3
    myForm.Label3.Font.Size = 11
    myForm.Label4 = strLabel4
    myForm.Label4.Font.Size = 13
    myForm.Label4.Font.Bold = True
    myForm.CommandButton1.Caption = button1
    myForm.CommandButton1.Font.Size = 11
    
    ' Help File
    myForm.Label5 = strLabel5
    myForm.Label5.Font.Size = 8
    myForm.Label5.Font.Italic = True
    
    Application.StatusBar = "Macro has been initiated."
    myForm.Show

End Sub
'---------------------------------------------------------------------------------------
' Date Created : August 13, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : August 13, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : HELPFILE
' Description  : This function will feed the help tip section depending on the button
'                that has been activated.
' Parameters   : String
' Returns      : String
'---------------------------------------------------------------------------------------
Function HELPFILE(ByVal Notification As Integer) As String

    Dim NotifyUser As String
    
    Select Case Notification
        Case 1
            NotifyUser = "TITLE: THE CLIMATE INDICES CONVERSION MACRO" & vbLf
            NotifyUser = NotifyUser & "DESCRIPTION: This macro will convert the " & _
                "climate indices result worksheet into a .CSV file. " & vbLf
            NotifyUser = NotifyUser & "INPUT: Find the location of all .XLSX files that needs to be converted." & vbLf
            NotifyUser = NotifyUser & "OUTPUT: .CSV files" & vbLf
    End Select
    
    HELPFILE = NotifyUser
    
End Function
'---------------------------------------------------------------------
' Date Created : January 9, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : January 10, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : SaveLogFile
' Description  : This function saves file as .TXT.
'                When new file is named after an existing file, the
'                same name is used with an number attached to it.
' Parameters   : String, String
' Returns      : -
'---------------------------------------------------------------------
Function SaveLogFile(ByVal fileDir As String, _
ByVal fileName As String, ByVal fileExt As String) As String

    Dim saveFile As String
    Dim formatDate As String
    Dim saveDate As String
    Dim saveName As String
    Dim sPath As String

    ' Date
    formatDate = Format(Date, "yyyy/mm/dd")
    saveDate = Replace(formatDate, "/", "")
    
    ' Save information as Temp, which can then be renamed later..
    sPath = fileDir
    If Right(fileDir, 1) <> "\" Then sPath = fileDir & "\"
    saveName = fileName & "_" & saveDate & fileExt
    
    ' Rename existing file
    i = 1
    If CheckFileExists(sPath, saveName) = True Then
        If Dir(sPath & fileName & "_" & saveDate & "_" & i & fileExt) <> "" Then
            Do Until Dir(sPath & fileName & "_" & saveDate & "_" & i & fileExt) = ""
                i = i + 1
            Loop
            saveFile = sPath & fileName & "_" & saveDate & "_" & i & fileExt
        Else: saveFile = sPath & fileName & "_" & saveDate & "_" & i & fileExt
        End If
    Else: saveFile = sPath & fileName & "_" & saveDate & fileExt
    End If
    
    SaveLogFile = saveFile
    
End Function


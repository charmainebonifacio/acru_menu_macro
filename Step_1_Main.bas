Attribute VB_Name = "Step_1_Main"
Public objFSOlog As Object
Public logfile As TextStream
Public logtxt As String
Public appSTATUS As String
'---------------------------------------------------------------------------------------
' Date Created : August 10, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : August 10, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ACRU_MENU_MAIN
' Description  : This is the main function that will (1) Save the latest worksheet into
'                a textfile and then (2) call the Fortran application using the new
'                file.
'---------------------------------------------------------------------------------------
Function ACRU_MENU_MAIN()

    Dim start_time As Date, end_time As Date
    Dim ProcessingTime As Long
    Dim MessageSummary As String, SummaryTitle As String
    Dim logfilename As String, logtextfile As String, logext As String
    Dim UserSelectedFolder As String
    Dim MAINFolder As String, MAINOUT As String
    Dim SavedFileStatus As Boolean
    Dim ParameterizationStatus As Boolean
    
    ' Initialize Variables
    SummaryTitle = "Zonal Statistics Macro Diagnostic Summary"
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    '---------------------------------------------------------------------
    ' I. FIND DIRECTORY
    '---------------------------------------------------------------------
    UserSelectedFolder = GetFolder
    Debug.Print UserSelectedFolder
    MAINFolder = ReturnFolderName(UserSelectedFolder)
    Debug.Print MAINFolder
    If Len(MAINFolder) = 0 Then GoTo Cancel
     
    '---------------------------------------------------------------------
    ' II. LOGFILE SETUP
    '---------------------------------------------------------------------
    MAINOUT = ReturnSubFolder(UserSelectedFolder, OUTDIR)   ' Location of folder
    Debug.Print MAINOUT
    CheckOUTFolder = CheckFolderExists(MAINOUT)

    logext = ".txt"
    logfilename = "acru_menu_macro_log"
    logtextfile = SaveLogFile(MAINOUT, logfilename, logext)
    
    Set objFSOlog = CreateObject("Scripting.FileSystemObject")
    Set logfile = objFSOlog.CreateTextFile(logtextfile, True)
    
    ' Maintain log starting from here
    logfile.WriteLine "[ START MACRO ] "
    logfile.WriteLine " "
    logfile.WriteLine "[ CALIBRATION SUMMARY] "
    logfile.WriteLine UserForm1.TextBox1.Value
    logfile.WriteLine " "
    logfile.WriteLine "Selected directory: " & UserSelectedFolder
    
    '---------------------------------------------------------------------
    ' III. START PROGRAM
    '---------------------------------------------------------------------
    Call WarningMessage
    start_time = Now()
    logfile.WriteLine "[ PROCESSING FILE SUMMARY ]"
    SavedFileStatus = PROCESSFILE(UserSelectedFolder)
    If SavedFileStatus = False Then GoTo Cancel
    logfile.WriteLine " "
    logfile.WriteLine "[ MENU PARAMETERIZATION SUMMARY ]"
    ParameterizationStatus = MENU_PARAMETERIZATION(UserSelectedFolder)
    If SavedFileStatus = False Then
        logfile.WriteLine "Could not successfully parameterize the MENU FILE. "
        logfile.WriteLine "Please check the log files. "
    End If
    
    '---------------------------------------------------------------------
    ' IV. END PROGRAM
    '---------------------------------------------------------------------
    end_time = Now()
    ProcessingTime = DateDiff("s", CDate(start_time), CDate(end_time))
    MessageSummary = MacroTimer(ProcessingTime)
    logfile.WriteLine " "
    logfile.WriteLine MessageSummary
    
Cancel:
    If Len(MAINFolder) = 0 Then
        MsgBox "No directory was selected."
    End If
    If SavedFileStatus = False Then
        end_time = Now()
        ProcessingTime = DateDiff("s", CDate(start_time), CDate(end_time))
        MessageSummary = MacroTimer(ProcessingTime)
        logfile.WriteLine "Could not file the right worksheet. No text file was saved."
    End If
    
    logfile.WriteLine " "
    logfile.WriteLine "[ END PROGRAM ] "

    ' Close Log File
    logfile.Close
    Set logfile = Nothing
    Set objFSOlog = Nothing
End Function

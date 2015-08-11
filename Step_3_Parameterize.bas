Attribute VB_Name = "Step_3_Parameterize"
'---------------------------------------------------------------------
' Date Created : August 10, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : August 10, 2015
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : MENU_PARAMETERIZATION
' Description  : This function finds the necesary directory and file
'                in order to create a batch file that will call the
'                fortran program 'Harmonic Analysis'. All .OUT files
'                will then be copied from the C:\Harmonic directory.
' Parameters   : String
' Returns      : Boolean
'---------------------------------------------------------------------
Function MENU_PARAMETERIZATION(ByVal fileDir As String) As Boolean

    Dim menuFile As String
    Dim menuparamFile As String
    Dim exeFile As String
    Dim menuFileExist As Boolean
    Dim menuparamFileExist As Boolean
    Dim exeFileExist As Boolean
    Dim ShellOpen As String

    MENU_PARAMETERIZATION = False
    
    ' The Following Files are required before the application can be called:
    ' MENU File
    menuFile = "MENU"
    menuFileExist = CheckFileExists(fileDir, menuFile)
    ' MENU_PARAM.txt
    menuparamFile = "MENU_PARAM.txt"
    menuparamFileExist = CheckFileExists(fileDir, menuparamFile)

    ' Check if Program File exists
    ' Program File
    exeFile = "menu_parameterization.exe"
    exeFileExist = CheckFileExists(fileDir, exeFile)
    
    If menuFileExist = True And menuparamFileExist = True And exeFileExist = True Then
        MENU_PARAMETERIZATION = True
        logtxt = "All required files were found within the directory."
        Debug.Print logtxt
        logfile.WriteLine " "
        logfile.WriteLine logtxt
    End If
    
    If MENU_PARAMETERIZATION = True Then
        Dim FSO As Object
        Dim BATFile As String, TargetFolderPath As String
        Dim ProcessID As Boolean, FileNumber As Integer
        ShellOpen = fileDir & "\" & exeFile
        BATFile = fileDir & "acru_menu.bat"
        logfile.WriteLine "Creating a Batch File: " & BATFile
        FileNumber = FreeFile() ' Get unused file number.
        Open BATFile For Output As #FileNumber ' Create file name.
        Print #FileNumber, "cd " & fileDir
        Print #FileNumber, exeFile
        Close #FileNumber ' Close file.
        ProcessID = Shell_AndWait(BATFile)
        Debug.Print ProcessID
        Kill BATFile ' Delete the file right away
        logfile.WriteLine "Deleted BAT File after Fortran program terminated."
    End If
End Function


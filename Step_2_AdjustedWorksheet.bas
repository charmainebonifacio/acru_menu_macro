Attribute VB_Name = "Step_2_AdjustedWorksheet"
'---------------------------------------------------------------------
' Date Created : August 10, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : August 10, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : PROCESSFILES
' Description  : This function will process one .XLSX file.
' Parameters   : String
' Returns      : String
'---------------------------------------------------------------------
Function PROCESSFILE(ByVal fileDir As String) As Boolean

    Dim objFolder As Object, objFSO As Object
    Dim wbSource As Workbook, SourceSheet As Worksheet
    Dim wbDest As Workbook, DestSheet As Worksheet
    Dim FileCounter As Long
    Dim sThisFilePath As String, sFile As String
    Dim GridName As String
    Dim VarType As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    appSTATUS = "Processed .txt files for " & VarSelected & "."
    Application.StatusBar = appSTATUS
   
    ' Status Bar Update
    appSTATUS = "Saving spreadsheet as a text file for fortran program...."
    Application.StatusBar = appSTATUS
    logtxt = appSTATUS
    logfile.WriteLine logtxt
    
    Dim openFile As Variant
    Dim fileCheck As Boolean
    Dim filePath As String

    '-------------------------------------------------------------
    ' Select Multiple ACRU Output files to be processed.
    '-------------------------------------------------------------
    openFile = Application.GetOpenFilename( _
        filefilter:="MENU PARAMETER FILE (*.xlsx*), *.xlsx*", _
        title:="Open Menu Parameter File", MultiSelect:=False)
    If TypeName(openFile) = "Boolean" Then GoTo Cancel '"User has cancelled."
            
    ' Open file and set it as source worksheet
    Set wbSource = Workbooks.Open(openFile)
    Set SourceSheet = wbSource.Worksheets(3)
    SourceSheet.Activate
    GridName = SourceSheet.Name
        
    logfile.WriteLine " "
    If GridName = "MENU_PARAM" Then
        PROCESSFILE = True
        logtxt = "Found the correct worksheet and will be save as a textfile."
        Debug.Print logtxt
        logfile.WriteLine logtxt
                
        ' Save Summary Workbook according to the variable type: RAD, SUN, REL, WND
        Call SaveTXT(ActiveWorkbook, SourceSheet, fileDir, GridName)
        logtxt = "Successfully created the text file to be used in the Fortran program."
        Debug.Print logtxt
        logfile.WriteLine logtxt
    Else
        PROCESSFILE = False
        logtxt = "The page couldn't be found."
        Debug.Print logtxt
        logfile.WriteLine logtxt
    End If
    wbSource.Close SaveChanges:=False
        
    If PROCESSFILES = False Then GoTo Cancel
    
Cancel:
    Set wbSource = Nothing
    Set SourceSheet = Nothing
    Set wbDest = Nothing
    Set DestSheet = Nothing
    Set objFSO = Nothing
    Set objFolder = Nothing
End Function


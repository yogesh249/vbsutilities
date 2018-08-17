'---------------------------------------------------------------------------------------
'
' Name:  FileCount.vbs
' Version: 1.0
' Date:  7-7-2009
' Author:  Yogesh Gandhi
' Description: FileCount.vbs calculates the Code files (.cc, .h, .c) files in subfolders within
'   a folder and puts the folder wise summary excel workbook
' 
' This can be used on our 9.1.0 spectrum folder, to check how many files are there
'   in which folder
'   This utility requires Microsoft Excel to be present on that machine.
'   If Microsoft Excel is not installed on that machine, you can get a tab seperated summary
'   Just uncomment/comment the relevant/irrelevant code 
' It will generate a filelist.txt file in the specified folder.
'---------------------------------------------------------------------------------------
FolderToBeAnalysed = BrowseFolder( "", False )
Call CountCCAndh(FolderToBeAnalysed, false)
'**********************************************************************
Sub CountCCAndh(sFolder, bIncludeDirInCount)
 Dim objFolder, objSubFolders, objFso, o, n 
 Set objFso = Createobject( "Scripting.FileSystemObject" )  
 Set objDialog = CreateObject( "SAFRCFileDlg.FileSave" )
 
 'Get the file name and path from the user.
 dim xlfilename
 ' Note: If no path is specified, the "current" directory will
 '       be the one remembered from the last "SAFRCFileDlg.FileOpen"
 '       or "SAFRCFileDlg.FileSave" dialog!
 objDialog.FileName = "test.xls"
 ' Note: The FileType property is cosmetic only, it doesn't
 '       automatically append the right file extension!
 '       So make sure you type the extension yourself!
 objDialog.FileType = "Excel file(*.xls)"
 If objDialog.OpenFileSaveDlg Then
  xlfilename = objDialog.FileName
 End If
 
 set objXL = CreateObject( "Excel.Application" )
 objXL.Visible = False
 objXL.WorkBooks.Add
 dim pcount
 dim objNewFile
 'Set objNewFile = objFso.CreateTextFile(sFolder & "\filelist.txt", True)  
 Set objFolder = objFso.GetFolder(sFolder) 
    ObjXL.ActiveSheet.Cells(1,1).Value = "Folder name"
    ObjXL.ActiveSheet.Cells(1,2).Value = "C/CC/h files"
    ObjXL.ActiveSheet.Cells(1,3).Value = "java files"
    ObjXL.ActiveSheet.Cells(1,4).Value = "JSP files"
    ObjXL.ActiveSheet.Cells(1,5).Value = "JS files"
    ObjXL.ActiveSheet.Cells(1,6).Value = "XML files"
    ObjXL.ActiveSheet.Cells(1,7).Value = "WSDL files"
    ObjXL.ActiveSheet.Cells(1,8).Value = "PL files"
    ObjXL.ActiveSheet.Cells(1,9).Value = "BAT files"
    ObjXL.ActiveSheet.Cells(1,10).Value = "SH files"
    ObjXL.ActiveSheet.Cells(1,11).Value = "RPT files"
    ObjXL.ActiveSheet.Cells(1,12).Value = "CUS files"
    ObjXL.ActiveSheet.Cells(1,13).Value = "CFG files"
    
 pcount = 2
 For Each o In objFolder.SubFolders 
 ' objNewFile.Write(o.Name)
 ' objNewFile.Write(vbTab)
 ' objNewFile.WriteLine(cntFiles(o, false))
  ObjXL.ActiveSheet.Cells(pcount,1).Value = o.Name
  ObjXL.ActiveSheet.Cells(pcount,2).Value = cntCFiles(o, false)
        ObjXL.ActiveSheet.Cells(pcount,3).Value = cntFiles(o, false, ".java")
        ObjXL.ActiveSheet.Cells(pcount,4).Value = cntFiles(o, false, ".jsp")
        ObjXL.ActiveSheet.Cells(pcount,5).Value = cntFiles(o, false, ".js")
        ObjXL.ActiveSheet.Cells(pcount,6).Value = cntFiles(o, false, ".xml")
        ObjXL.ActiveSheet.Cells(pcount,7).Value = cntFiles(o, false, ".wsdl")
        ObjXL.ActiveSheet.Cells(pcount,8).Value = cntFiles(o, false, ".pl")
        ObjXL.ActiveSheet.Cells(pcount,9).Value = cntFiles(o, false, ".bat")
        ObjXL.ActiveSheet.Cells(pcount,10).Value = cntFiles(o, false, ".sh")
        ObjXL.ActiveSheet.Cells(pcount,11).Value = cntFiles(o, false, ".rpt")
        ObjXL.ActiveSheet.Cells(pcount,12).Value = cntFiles(o, false, ".cus")
        ObjXL.ActiveSheet.Cells(pcount,13).Value = cntFiles(o, false, ".cfg")
        
  pcount = pcount + 1
 Next 
 
 'Finally close the workbook
   ObjXL.ActiveWorkbook.SaveAs(xlfilename)
   ObjXL.Application.Quit
   Set ObjXL = Nothing 
   msgbox xlfilename & " Saved"
End Sub
'***********************************************************************************
Function cntFiles( strFolder, bIncludeDirInCount, ext ) 
 Dim objFolder, objSubFolders, objFso, o, n
 On Error Resume Next
 cntFiles = -1
 Set objFso = Createobject( "Scripting.FileSystemObject" ) 
 Set objFolder = objFso.GetFolder(strFolder) 
 If( Err.Number <> 0 ) Then 
 Exit Function 
 End If 
 'n = objFolder.files.count 
 for each fl in ObjFolder.files
  if Right(fl.name, len(ext))=ext then
   n = n + 1
  end if
 next
 Set objSubFolders = objFolder.SubFolders 
 For Each o In objSubFolders 
  n = n + cntFiles( o, bIncludeDirInCount, ext ) 
  If( bIncludeDirInCount ) Then 
   n = n + 1 
  End If 
 Next
 Set objSubFolders = Nothing 
 Set objFolder = Nothing
 cntFiles = n 
End Function
'*************************************************************************************************
Function cntCFiles( strFolder, bIncludeDirInCount ) 
 Dim objFolder, objSubFolders, objFso, o, n
 On Error Resume Next
 cntCFiles = -1
 Set objFso = Createobject( "Scripting.FileSystemObject" ) 
 Set objFolder = objFso.GetFolder(strFolder) 
 If( Err.Number <> 0 ) Then 
 Exit Function 
 End If 
 'n = objFolder.files.count 
 for each fl in ObjFolder.files
  if Right(fl.name, 3)=".cc" or Right(fl.name,2) = ".h" or Right(fl.name, 2) = ".c" then
   n = n + 1
  end if
 next
 Set objSubFolders = objFolder.SubFolders 
 For Each o In objSubFolders 
  n = n + cntCFiles( o, bIncludeDirInCount ) 
  If( bIncludeDirInCount ) Then 
   n = n + 1 
  End If 
 Next
 Set objSubFolders = Nothing 
 Set objFolder = Nothing
 cntCFiles = n 
End Function
'**********************************************************************************************************
Function cntJavaFiles( strFolder, bIncludeDirInCount ) 
 Dim objFolder, objSubFolders, objFso, o, n
 On Error Resume Next
 cntJavaFiles = -1
 Set objFso = Createobject( "Scripting.FileSystemObject" ) 
 Set objFolder = objFso.GetFolder(strFolder) 
 If( Err.Number <> 0 ) Then 
 Exit Function 
 End If 
 'n = objFolder.files.count 
 for each fl in ObjFolder.files
  if Right(fl.name, 5)=".java" then
   n = n + 1
  end if
 next
 Set objSubFolders = objFolder.SubFolders 
 For Each o In objSubFolders 
  n = n + cntJavaFiles( o, bIncludeDirInCount ) 
  If( bIncludeDirInCount ) Then 
   n = n + 1 
  End If 
 Next
 Set objSubFolders = Nothing 
 Set objFolder = Nothing
 cntJavaFiles = n 
End Function
'***********************************************************************************************
Function BrowseFolder( myStartLocation, blnSimpleDialog )
' This function generates a Browse Folder dialog
' and returns the selected folder as a string.
'
' Arguments:
' myStartLocation   [string]  start folder for dialog, or "My Computer", or
'                             empty string to open in "Desktop\My Documents"
' blnSimpleDialog   [boolean] if False, an additional text field will be
'                             displayed where the folder can be selected
'                             by typing the fully qualified path
'
' Returns:          [string]  the fully qualified path to the selected folder
'
' Based on the Hey Scripting Guys article
' "How Can I Show Users a Dialog Box That Only Lets Them Select Folders?"
' http://www.microsoft.com/technet/scriptcenter/resources/qanda/jun05/hey0617.mspx
'
' Function written by Rob van der Woude
' http://www.robvanderwoude.com
    Const MY_COMPUTER   = &H11&
    Const WINDOW_HANDLE = 0 ' Must ALWAYS be 0
    Dim numOptions, objFolder, objFolderItem
    Dim objPath, objShell, strPath, strPrompt
    ' Set the options for the dialog window
    strPrompt = "Select a folder:"
    If blnSimpleDialog = True Then
        numOptions = 0      ' Simple dialog
    Else
        numOptions = &H10&  ' Additional text field to type folder path
    End If
    
    ' Create a Windows Shell object
    Set objShell = CreateObject( "Shell.Application" )
    ' If specified, convert "My Computer" to a valid
    ' path for the Windows Shell's BrowseFolder method
    If UCase( myStartLocation ) = "MY COMPUTER" Then
        Set objFolder = objShell.Namespace( MY_COMPUTER )
        Set objFolderItem = objFolder.Self
        strPath = objFolderItem.Path
    Else
        strPath = myStartLocation
    End If
    Set objFolder = objShell.BrowseForFolder( WINDOW_HANDLE, strPrompt, _
                                              numOptions, strPath )
    ' Quit if no folder was selected
    If objFolder Is Nothing Then
        BrowseFolder = ""
        Exit Function
    End If
    ' Retrieve the path of the selected folder
    Set objFolderItem = objFolder.Self
    objPath = objFolderItem.Path
    ' Return the path of the selected folder
    BrowseFolder = objPath
End Function 
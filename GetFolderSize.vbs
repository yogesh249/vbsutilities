'---------------------------------------------------------------------------------------
'
' Name:  getfoldersize.vbs
' Version: 1.0
' Date:  7-5-2002
' Author:  Hans van der Zaag
' Description: getfoldersize.vbs calculates the size of all subfolders within
'   a folder and sorts this data in an excel workbook
' 
'---------------------------------------------------------------------------------------
   'rootfolder = Inputbox("Enter directory/foldername: " & _
'                         chr(10) & chr(10) & "(i.e. C:\Program Files or " & _
 '                        "\\Servername\C$\Program Files)" & chr(10) & chr(10), _
  '                       "Getfoldersize", "C:\Program Files")
   outputfile = "C:\foldersize_" & Day(now) & Month(now) & Year(now) & ".xls"
 
 'Parse the passed parameters.
 const FIRST_ITEM=0
 dim LAST_ITEM
 dim args_index
 dim command_line_args
 set command_line_args = wscript.Arguments
 'Check if any Arguments have been passed to our script.
 'rootfolder=command_line_args(FIRST_ITEM)
    rootfolder = BrowseFolder( "", False )
   Set fso = CreateObject("scripting.filesystemobject")
   if fso.fileexists(outputfile) then fso.deletefile(outputfile)
'Create Excel workbook
   set objXL = CreateObject( "Excel.Application" )
   objXL.Visible = False
   objXL.WorkBooks.Add
'Counter 1 for writing in cell A1 within the excel workbook
   icount = 1
'Run checkfolder
   CheckFolder (FSO.getfolder(rootfolder))
Sub CheckFolder(objCurrentFolder)
       For Each objFolder In objCurrentFolder.SubFolders
         FolderSize = objFolder.Size
         Tmp = (FormatNumber(FolderSize, 0, , , 0)/1024)/1024
         ObjXL.ActiveSheet.Cells(icount,1).Value = objFolder.Path
         ObjXL.ActiveSheet.Cells(icount,2).Value = Tmp
         'Wscript.Echo Tmp & " " & objFolder.Path
  'raise counter with 1 for a new row in excel 
         icount = icount + 1
       Next
       'Recurse through all of the folders
'       For Each objNewFolder In objCurrentFolder.subFolders
'               CheckFolder objNewFolder
'       Next


   rng = "A6:B" & (6+icount-2)
   ObjXL.Range(rng).Select
   ObjXL.ActiveSheet.Shapes.AddChart.Select
   ObjXL.ActiveChart.SetSourceData ObjXL.ActiveSheet.Range(rng)
   ' The constant defined in VBA is xlPie, but here it is not working
   ' so hard-coding the value here.
   ObjXL.ActiveChart.ChartType = 5

       
End Sub
if icount=1 then 
 MsgBox "No Subfolders are present in this folder", vbInformation, "No subfolders"
 WScript.Quit
end if
'sort data in excel
objXL.ActiveCell.CurrentRegion.Select
objXL.Selection.Sort objXL.Worksheets(1).Range("B1"), _
                   2, _
                   , _
                   , _
                   , _
                   , _
                   , _
                   0, _
                   1, _
                   False, _
                   1
'Lay out for Excel workbook 
   objXL.Range("A1").Select
   objXL.Selection.EntireRow.Insert
   objXL.Selection.EntireRow.Insert
   objXL.Selection.EntireRow.Insert
   objXL.Selection.EntireRow.Insert
   objXL.Selection.EntireRow.Insert
   objXL.Columns(1).ColumnWidth = 60
   objXL.Columns(2).ColumnWidth = 15
   objXL.Columns(2).NumberFormat = "#,##0.0"
   objXL.Range("B1:B1").NumberFormat = "d-m-yyyy"
   objXL.Range("A1:B5").Select
   objXL.Selection.Font.Bold = True
   objXL.Range("A1:B3").Select
   objXL.Selection.Font.ColorIndex = 5
   objXL.Range("A1:A1").Select
   objXL.Selection.Font.Italic = True
   objXL.Selection.Font.Size = 16
   ObjXL.ActiveSheet.Cells(1,1).Value = "Survey FolderSize " 
   ObjXL.ActiveSheet.Cells(1,2).Value = Day(now) & "-" & Month(now) & "-"& Year(now)
   ObjXL.ActiveSheet.Cells(3,1).Value = UCase(rootfolder)
   ObjXL.ActiveSheet.Cells(5,1).Value = "Folder"
   ObjXL.ActiveSheet.Cells(5,2).Value = "Total (MB)"

'Finally close the workbook
   ObjXL.ActiveWorkbook.SaveAs(outputfile)
   ObjXL.Application.Quit
   Set ObjXL = Nothing
'Message when finished
   Set WshShell = CreateObject("WScript.Shell")
   Finished = Msgbox ("Script executed successfully, results can be found in " & Chr(10) _
                     & outputfile & "." & Chr(10) & Chr(10) _
                     & "Do you want to view the results now?", 65, "Script executed successfully!")
   if Finished = 1 then WshShell.Run "excel " & outputfile
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


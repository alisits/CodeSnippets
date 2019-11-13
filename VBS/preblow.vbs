'Script for converting data for BlowMould Simulator
'Creates a mesh for a parison using imput machine data from a .txt file
'Created by Alicia Garcia, algarcia@de.tiauto.com

On error resume next

Dim xlApp, xlBook, ft1, newname, newname2, myfile, currentdir, objExcel, objFso, objFolder, objFile, excelfilename , xlname
Dim proc 

'Name and path of the Excel Data Conversor (where preblow.vbs is)
Set fso = CreateObject("Scripting.FileSystemObject")
currentdir=fso.getabsolutepathname(".")
'Name of the excel file with the macro
excelfilename=fso.buildpath(currentdir, "preblow.xlsm") 'Introduce here right name for script
griddir=currentdir & "\Grid\" 'Subfolder Grid

'Deletes all previous .xls files in Grid
Set obj = CreateObject("Scripting.FileSystemObject")
obj.deletefile(griddir & "*.xls")

'Runs excel file where macro is stored
Set xlApp = CreateObject("Excel.Application")
xlApp.Application.Visible=False
set xlbook= xlApp.Workbooks.Open (excelfilename)

'Orders to run the excel macros
Wscript.echo "Working.. please wait, it takes a while. Press ESC to cancel."
xlApp.Run "Macro1" 'This macro saves as xls the edited File
xlbook.close (false)
xlApp.Quit
set xlbook =Nothing
set xlApp = Nothing


'Reads ANY file with xls extension (there should only be the one we created as the others were deleted previously),
'and saves executes the function that saves as .bcs
    Set objfso = CreateObject("Scripting.FileSystemObject")
     objStartFolder = griddir 
     Set objFolder = objfso.GetFolder(objStartFolder)
     For Each objFile In objFolder.Files
          If objfso.GetExtensionName(objFile) = "xls" Then	
              ExcelConvert (objFile.name)
          End If
     Next

'Function for reading the xls file and saving as a .bcs
Sub ExcelConvert(SourceFile) 
	Dim AppExcel 
	Dim OpenWorkbook 
	dim newname, wholename
	Const xlS = 1 
	Set objShell = CreateObject("Wscript.Shell")
	objShell.CurrentDirectory = griddir 'changes Directory to Grid
	Set AppExcel = CreateObject("Excel.Application") 
	AppExcel.Visible = False
	wholename=griddir&"\"&sourcefile 'gets the name
	Set OpenWorkbook = AppExcel.Workbooks.Open(wholename) 
	newname = fso.BuildPath(openworkbook.path, fso.GetBaseName(openworkbook.Name) & ".bcs") 
	openworkBook.SaveAs newname, -4158 'saves as bcs with same name

	OpenWorkbook.Close false
	AppExcel.Quit 
	Set OpenWorkbook = Nothing
	Set AppExcel = Nothing 
End Sub 

System.Runtime.InteropServices.Marshal.ReleaseComObject(Application)


For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
proc.Kill()
Next

WScript.Echo "Finished."
WScript.Quit
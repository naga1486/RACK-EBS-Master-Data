'General Header
'#####################################################################################################################
'Script Description		: Init Script
'Test Tool/Version		: HP Quick Test Professional 9.5 and above
'Test Tool Settings		: N.A.
'Application Automated	: N.A.
'Author					: Rackspace
'Date Created			: 02/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Option Explicit	'Forcing Variable declarations

Dim gstrRelativePath, garrQtpAddins, gintFwPosition, gstrTimeStamp
Dim gobjTimeStampFolder
Dim gobjQtpFolder, gobjExcelFolder, gobjHtmlFolder, gobjScreenshotsFolder
Dim gobjFso, gobjMyFile
Dim gobjQtpApp

Dim SheetVal 'Variable to get the select ed  EBS module name from the POPUP

Set gobjFso = CreateObject("Scripting.FileSystemObject")
Set gobjQtpApp = CreateObject("QuickTest.Application")

'Store the relative path 
gstrRelativePath = gobjFso.GetParentFolderName(WScript.ScriptFullName)

'Initialise the array of QTP add-ins
garrQtpAddins = Array("ActiveX","Visual Basic","Java","Web","Oracle")

'Close QTP if already open
If gobjQtpApp.Launched Then
	gobjQtpApp.Quit
End If

'Load required add-ins in QTP
Dim gblnActivateOK, gstrError
gblnActivateOK = gobjQtpApp.SetActiveAddins(garrQtpAddins, gstrError)
If Not gblnActivateOK Then	'If a problem occurs while loading the add-ins
	MsgBox gstrError	'Show a message containing the error
	WScript.Quit	'Terminate the init script
End If

'Open QTP with the required add-ins loaded
gobjQtpApp.Launch
gobjQtpApp.Visible = True

'Set general QTP options as required
RACK_SetQtpOptions()

'Add the relative path to QTP's folders list (for easy portability of scripts)
gintFwPosition = gobjQtpApp.Folders.Find(gstrRelativePath)
If gintFwPosition <> -1 Then	'If the folder is already found in the collection
	gobjQtpApp.Folders.Remove gintFwPosition
End If
gobjQtpApp.Folders.Add gstrRelativePath, 1	'Add the folder to the collection in position 1

'Initialise StopAllExecution temp file
Set gobjMyFile = gobjFso.CreateTextFile(gstrRelativePath & "\StopAllExecution.txt", True)
gobjMyFile.Close
Set gobjMyFile = gobjFso.OpenTextFile(gstrRelativePath &_
											"\StopAllExecution.txt", 2)	'Open the StopAllExecution file for writing
gobjMyFile.Writeline("False")
gobjMyFile.Close

'Create Results folder with timestamp
gstrTimeStamp = "Run" & "_" & Replace(Date(),"/","-") & "_" &_
												Replace(Time(),":","-")

Set gobjTimeStampFolder = gobjFso.CreateFolder(gstrRelativePath &_
											"\Results\" & gstrTimeStamp)
Set gobjExcelFolder	= _
				gobjFso.CreateFolder(gobjTimeStampFolder & "\Excel Results")

Set gobjHtmlFolder = _
				gobjFso.CreateFolder(gobjTimeStampFolder & "\HTML Results")
Set gobjQtpFolder = _
				gobjFso.CreateFolder(gobjTimeStampFolder & "\QTP Results")
Set gobjScreenshotsFolder = _
				gobjFso.CreateFolder(gobjTimeStampFolder & "\Screenshots")

'Switch off Debug mode if On
Dim gblnDebugMode
gblnDebugMode = RACK_GetConfig("DebugMode")
If (gblnDebugMode = "True") Then
	RACK_SetConfig "DebugMode", "False"
End If

set SheetVal = Nothing
'################################### POPUP WINDOW FUNCTIONALITY ##################################################################################

'Calling Popup Subroutiine
ExcelMacroPopUp()

'Popup Subrotine Creation
Sub ExcelMacroPopUp() 

  Dim xlApp 
  Dim xlBook 

  Set xlApp = CreateObject("Excel.Application") 
  Set xlBook = xlApp.Workbooks.Open(gstrRelativePath &"\Datatables\RackSpace_Macro.xlsm") 
  xlApp.Run "RS"
  xlApp.Quit 

  Set xlBook = Nothing 
  Set xlApp = Nothing 
 
  Dim objExcel, objWorkBook, objWorkSheet
  Dim intRowCount, intRowIterator
  Set objExcel = CreateObject("Excel.Application")

  Set objWorkBook = objExcel.Workbooks.Open(gstrRelativePath & "\Datatables\RackSpace_Macro.xlsm")

	'Set objWorkBook = objExcel.Workbooks.Open("D:\Rackspace\RACK - EBS Master-Data-New-12-01-2015\RACK - EBS Master-Data-New-12-01-2015\TEST DATA\Rack_Variable.xlsx")

  Set objWorkSheet = objWorkBook.WorkSheets("Sheet1") 
  objWorkSheet.Activate
  SheetVal = objWorkSheet.Range("A" & "1").Value

    objWorkBook.Save
	objWorkBook.Close
	
	objExcel.quit
	
	'Release all objects
	Set objWorkSheet = Nothing
	Set objWorkBook = Nothing
	Set objExcel = Nothing

  

' Calling Batch Run Subroutine
'RACK_BatchRun(SheetVal)

End Sub

'################################## END OF POPUP WINDOW FUNCTIONALITY #######################################################################

'Call the function for batch execution of test cases
Call RACK_BatchRun(SheetVal)

'Export a copy of the Test data sheets
gobjFso.CopyFolder gstrRelativePath & "\Datatables", gobjTimeStampFolder & "\Datatables", True
Dim gstrSvnFolder
gstrSVNFolder = gobjTimeStampFolder & "\Datatables\.svn"
If gobjFso.FolderExists(gstrSvnFolder) Then
	gobjFso.GetFolder(gstrSvnFolder).Delete(True)
End If
'Restore original Debug mode setting
RACK_SetConfig "DebugMode", gblnDebugMode

'Delete StopAllExecution temp file
If (gobjFso.FileExists(gstrRelativePath & "\StopAllExecution.txt")) Then
	gobjFso.DeleteFile(gstrRelativePath & "\StopAllExecution.txt")
End If

'Close QTP
'gobjQtpApp.Quit

'Display HTML results at the end of Test Run
Dim gobjShell 
Set gobjShell = CreateObject("WScript.Shell")
'msgbox gobjHtmlFolder
gobjShell.Run """" & gobjHtmlFolder & "\Summary.html"""

'Release all objects
Set gobjFso = Nothing
Set gobjMyFile = Nothing
Set gobjTimeStampFolder = Nothing
Set gobjExcelFolder = Nothing
Set gobjHtmlFolder = Nothing
Set gobjQtpFolder = Nothing
Set gobjScreenshotsFolder = Nothing
Set gobjQtpApp = Nothing
Set gobjShell = Nothing
'#####################################################################################################################


'#####################################################################################################################
'Function Description	: Function to get the configuration data from the RACK.ini configuration file
'Input Parameters 		: strKey
'Return Value    		: Corresponding value from RACK.ini
'Author					: Rackspace
'Date Created			: 02/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Function RACK_GetConfig(strKey)
	Dim strLine, arrLine, strValue, strConfigFilePath
	strConfigFilePath = gstrRelativePath & "\RACK.ini"
	Set gobjMyFile = gobjFso.OpenTextFile(strConfigFilePath,1)
	Do Until gobjMyFile.AtEndOfStream
		strLine = gobjMyFile.ReadLine
		If strLine <> "" Then
			arrLine = split(strLine,"=")
			If arrLine(0) = strKey Then
				strValue = arrLine(1)
				Exit Do
			End If
		End If
	Loop
	
	gobjMyFile.close()
	RACK_GetConfig = CStr(strValue)
End Function
'#####################################################################################################################

'#####################################################################################################################
'Function Description	: Function to set a particular value for the configuration data in the RACK.ini configuration file
'Input Parameters 		: strKey, strValue
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 02/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_SetConfig(strKey, strValue)
	Dim objNewFile, objOldFile, strLine, arrLine, strOldPath, strNewPath
	strNewPath = gstrRelativePath & "\RACK1.ini"
	strOldPath = gstrRelativePath & "\RACK.ini"
	Set objNewFile = gobjFso.OpenTextFile(strNewPath, 2, True)
	Set objOldFile = gobjFso.OpenTextFile(strOldPath, 1)
	
	Do Until objOldFile.AtEndOfStream
		strLine = objOldFile.ReadLine
		If strLine <> "" Then
			arrLine = split(strLine,"=")
			If arrLine(0) = strKey Then			
				objNewFile.WriteLine strKey & "=" & strValue
			Else
				objNewFile.WriteLine strLine
			End If
		End If
	Loop
	
	objNewFile.close()
	objOldFile.close()
	gobjFso.DeleteFile strOldPath
	gobjFso.MoveFile strNewPath, strOldPath
	
	'Release all objects
	Set objNewFile = Nothing
	Set objOldFile = Nothing
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description	: Function to set general QTP options as required
'Input Parameters 		: None
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 02/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_SetQtpOptions()
	gobjQtpApp.Options.Run.ViewResults = False	
	'gobjQtpApp.Options.Run.ImageCaptureForTestResults = "OnError"
	'gobjQtpApp.Options.Run.MovieCaptureForTestResults = "Never"
	'gobjQtpApp.Options.Run.RunMode = "Fast"
End Sub
'#####################################################################################################################


'#####################################################################################################################
'Function Description	: Function to execute the Test Batch Run
'Input Parameters 		: None
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 02/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_BatchRun(strFileName)

	Dim intIterator, arrScenarios, arrTestCases
	'RACK_GetRunInfo "Datatables\RackSpace_Macro.xls","Start ", arrScenarios
	RACK_LoadDriverScript()

	RACK_GetRunInfo "Datatables\" & strFileName  & ".xls" , _
				strFileName, arrTestCases
	RACK_Execute strFileName, arrTestCases

'	For intIterator = 1 to UBOUND(arrScenarios,1)-1
'		If strcomp(arrScenarios (intIterator,3),"True") = 0 then
'					RACK_GetRunInfo "Datatables\" & arrScenarios (intIterator,1)  & ".xls" , _
'								arrScenarios (intIterator,1), arrTestCases
'					RACK_Execute arrScenarios (intIterator,1), arrTestCases
'				End If	
'	Next
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description	: Function to collect the data from sheet in the array
'Input Parameters 		: strWorkBook, strSheetName, arrRunInfo
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 02/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_GetRunInfo (strWorkBook, strSheetName, arrRunInfo)
	Dim objExcel, objWorkBook, objWorkSheet
	Dim intRowCount, intRowIterator
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkBook = objExcel.Workbooks.Open(gstrRelativePath & "\" & strWorkBook)
	Set objWorkSheet = objWorkBook.WorkSheets("" & strSheetName & "") 
	objWorkSheet.Activate
	intRowCount = objExcel.ActiveSheet.UsedRange.Rows.Count
	
	ReDim arrRunInfo(intRowCount,6)
	For intRowIterator = 2 to intRowCount
		arrRunInfo(intRowIterator -1,1) = objWorkSheet.Range("A" & intRowIterator).Value	'Scenario Name/Test Case ID
		arrRunInfo(intRowIterator -1,2) = objWorkSheet.Range("B" & intRowIterator).Value	'Description
		arrRunInfo(intRowIterator -1,3) = objWorkSheet.Range("C" & intRowIterator).Value	'Execute Flag
		arrRunInfo(intRowIterator -1,4) = objWorkSheet.Range("D" & intRowIterator).Value	'Iteration Mode
		arrRunInfo(intRowIterator -1,5) = objWorkSheet.Range("E" & intRowIterator).Value	'Start Iteration
		arrRunInfo(intRowIterator -1,6) = objWorkSheet.Range("F" & intRowIterator).Value	'End Iteration
	Next
	objWorkBook.Save
	objWorkBook.Close
	'objWorkBook.Save
	objExcel.quit
	
	'Release all objects
	Set objWorkSheet = Nothing
	Set objWorkBook = Nothing
	Set objExcel = Nothing
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description	: Function to load the Driver script and associate relevant artifacts
'Input Parameters 		: None
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 02/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_LoadDriverScript()	
	gobjQtpApp.Open gstrRelativePath & "\Driver Script"	
	
	'Note: The following sections are optional and can be commented out if done already
	
	'Associate relevant add-ins
	RACK_AssociateAddins()
	
	'Associate relevant object repositories
	RACK_AssociateRepositories()
	
	'Associate relevant libraries
	RACK_AssociateLibraries()
	
	'Associate relevant recovery scenarios
	'RACK_AssociateRecoveryScenarios()
	
	'Save the Driver script
	gobjQtpApp.Test.Save
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description	: Function to associate all the required QTP add-ins
'Input Parameters 		: None
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 02/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_AssociateAddins()
	Dim blnAddinsAssociated, strError
	blnAddinsAssociated = gobjQtpApp.Test.SetAssociatedAddins (garrQtpAddins, strError)
	If Not blnAddinsAssociated Then	'If a problem occurs while associating the add-ins
		MsgBox strError	'Show a message containing the error
		WScript.Quit	'Terminate the init script
	End If
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description	: Function to associate all the OR's from the OR folder
'Input Parameters 		: None
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 02/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_AssociateRepositories()
	Dim objRepositoryFolder, objRepositoryFileList, objQtpRepositories, objFile
	Set objRepositoryFolder = gobjFso.GetFolder(gstrRelativePath & "\Business Components\Libraries\Object Repository")
	Set objRepositoryFileList = objRepositoryFolder.Files
	Set objQtpRepositories = gobjQtpApp.Test.Actions("Action1").ObjectRepositories
	
	objQtpRepositories.RemoveAll
	For each objFile in objRepositoryFileList
		If Right(Ucase(objFile.Path),Len("TSR")+1) = ".TSR" Then
			objQtpRepositories.Add "Business Components\Libraries\Object Repository\" & objFile.Name
		End If
	Next
	
	'Release all objects
	Set objRepositoryFolder = Nothing
	Set objFile = Nothing
	Set objRepositoryFileList = Nothing
	Set objQtpRepositories = Nothing
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description	: Function to associate all the SL, RL and Business Components
'Input Parameters 		: None
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 02/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_AssociateLibraries()
	Dim objLibraryFolder, objLibraryFileList, objQtpSettings, objQtpLibraries, objFile
	Set objQtpSettings = gobjQtpApp.Test.Settings
	Set objQtpLibraries = objQtpSettings.Resources.Libraries
	objQtpLibraries.RemoveAll
	
	'Associate all .vbs files in SL folder
	Set objLibraryFolder = gobjFso.GetFolder(gstrRelativePath & "\Business Components\Libraries\Support Libraries")
	Set objLibraryFileList = objLibraryFolder.Files
	
	For each objFile in objLibraryFileList
		If Right(Ucase(objFile.Path),Len("VBS")+1) = ".VBS" Then
			objQtpLibraries.Add "Business Components\Libraries\Support Libraries\" & objFile.Name
		End If
	Next
	
	'Associate all .vbs files in RL folder
	Set objLibraryFolder = gobjFso.GetFolder(gstrRelativePath & "\Business Components\Libraries\Recovery Libraries")
	Set objLibraryFileList = objLibraryFolder.Files
	
	For each objFile in objLibraryFileList
		If Right(Ucase(objFile.Path),Len("VBS")+1) = ".VBS" Then
			objQtpLibraries.Add "Business Components\Libraries\Recovery Libraries\" & objFile.Name
		End If
	Next
	
	'Associate all .vbs files in Business Components folder
	Set objLibraryFolder = gobjFso.GetFolder(gstrRelativePath & "\Business Components")
	Set objLibraryFileList = objLibraryFolder.Files
	
	For each objFile in objLibraryFileList
		If Right(Ucase(objFile.Path),Len("VBS")+1) = ".VBS" Then
			objQtpLibraries.Add "Business Components\" & objFile.Name
		End If
	Next
	
	'Release all objects
	Set objLibraryFolder = Nothing
	Set objFile = Nothing
	Set objLibraryFileList = Nothing
	Set objQtpLibraries = Nothing
	Set objQtpSettings = Nothing
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description	: Function to associate the Recovery Scenarios
'Input Parameters 		: None
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 02/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_AssociateRecoveryScenarios()
	Dim objQtpSettings, objQtpTestRecovery
	Set objQtpSettings = gobjQtpApp.Test.Settings
	Set objQtpTestRecovery = objQtpSettings.Recovery	
	objQtpTestRecovery.RemoveAll
	
	'Associate required recovery scenarios
	objQtpTestRecovery.Add "Recovery Libraries\RACK_Recovery.qrs", "ObjNotFound"
	objQtpTestRecovery.Add "Recovery Libraries\RACK_Recovery.qrs", "Any Error"
	objQtpTestRecovery.Enabled = True
	
	'Release all objects
	Set objQtpTestRecovery = Nothing
	Set objQtpSettings = Nothing
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description	: Function to execute the test cases based on the Run Manager configuration
'Input Parameters 		: strTestScenario, arrTestCases
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 02/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_Execute(strTestScenario, arrTestCases)
	Dim intIterator, blnStopAllExecution
	For intIterator = 1 to UBOUND(arrTestCases,1) - 1
		Set gobjMyFile = gobjFso.OpenTextFile(gstrRelativePath & "\StopAllExecution.txt", 1)	'Open the StopAllExecution file for reading
		blnStopAllExecution = CBool(gobjMyFile.Readline())
		gobjMyFile.Close
		If (strcomp(arrTestCases (intIterator,3),"True") = 0 And blnStopAllExecution <> True) Then
			RACK_RunTest strTestScenario, arrTestCases(intIterator,1), _
					arrTestCases(intIterator,2), arrTestCases(intIterator,4), _
					arrTestCases(intIterator,5), arrTestCases(intIterator,6)
		End If
	Next
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description	: Function to run the Driver Script passing the test case ID
'Input Parameters 		: strTestScenario, strTestCaseID, strDescription, strIterationMode, intStartIteration, intEndIteration
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 02/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_RunTest(strTestScenario, strTestCaseID, strDescription, strIterationMode, intStartIteration, intEndIteration)
	Dim objQtpParamDefns	'As QuickTest.ParameterDefinitions
	Dim objQtpParams	'As QuickTest.Parameters
	Dim objQtpParam	'As QuickTest.Parameter
	
	Set objQtpParamDefns = gobjQtpApp.Test.ParameterDefinitions
	Set objQtpParams = objQtpParamDefns.GetParameters()
	
	Set objQtpParam = objQtpParams.Item("CurrentScenario")
	objQtpParam.Value = strTestScenario
	Set objQtpParam = objQtpParams.Item("CurrentTestCase")
	objQtpParam.Value = strTestCaseID
	Set objQtpParam = objQtpParams.Item("TestCaseDescription")
	objQtpParam.Value = strDescription
	Set objQtpParam = objQtpParams.Item("TimeStamp")
	objQtpParam.Value = gstrTimeStamp
	Set objQtpParam = objQtpParams.Item("IterationMode")
	Select Case strIterationMode
		Case "Run one iteration only"
			strIterationMode = "oneIteration"
		Case "Run all iterations"
			strIterationMode = "rngAll"
		Case "Run from <Start Iteration> to <End Iteration>"
			strIterationMode = "rngIterations"
	End Select
	objQtpParam.Value = strIterationMode
	Set objQtpParam = objQtpParams.Item("StartIteration")
	objQtpParam.Value = intStartIteration
	Set objQtpParam = objQtpParams.Item("EndIteration")
	objQtpParam.Value = intEndIteration
	
	'Create a separate folder for results of each test case
	Dim objResultsFolder
	Set objResultsFolder = gobjFso.CreateFolder(gobjQtpFolder &_
													"\" & strTestCaseID)	
	
	'Run the test with changed results options and parameters
	Dim objQtpResultsOpt
	Set objQtpResultsOpt = CreateObject("QuickTest.RunResultsOptions")
	objQtpResultsOpt.ResultsLocation = objResultsFolder
	gobjQtpApp.Test.Run objQtpResultsOpt, True, objQtpParams
	
	'Release all objects
	Set objQtpParamDefns = Nothing
	Set objQtpParams = Nothing
	Set objQtpParam = Nothing
	Set objResultsFolder = Nothing
	Set objQtpResultsOpt = Nothing
End Sub
'#####################################################################################################################




'#####################################################################################################################

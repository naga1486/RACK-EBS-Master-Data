'General Header
'#####################################################################################################################
'Script Description		: RACK Error handling functions
'Test Tool/Version		: HP Quick Test Professional 9.5 and above
'Test Tool Settings		: N.A.
'Application Automated	: Flight Application
'Author					: Rackspace
'Date Created			: 05/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Option Explicit	'Forcing Variable declarations

'#####################################################################################################################
'Function Description   : Function to perform error handling for the framework
						  ' This function can be customized to automatically logout and close the application on errors, handle unexpected popups, etc. 
'Input Parameters 		: None
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 05/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_ErrHandler()
	If (Err.Number <> 0) Then
		Select Case Environment.Value("OnError")
			'Stop & Dialog options are not relevant when run from QC Test Lab
			Case "NextStep"
				RACK_ReportEvent "RACK_Info",_
					"Error occurred during execution! Error description"&_
					"given below. Refer QTP Results for full details..", "Fail"
				RACK_LogError False
			Case "NextIteration"
				RACK_LogError False
				Environment.Value("ExitIteration") = True	'Exit current test case iteration
				General_CloseFlightApp()
			Case "NextTestCase"
				RACK_LogError False
				Environment.Value("StopExecution") = True	'Stop current test case execution
			Case "Stop"
				RACK_LogError True
				Environment.Value("StopExecution") = True	'Stop current test case execution
			Case "Dialog"
				MsgBox Err.Description
				RACK_LogError True
				'ExitRun
				Environment.Value("StopExecution") = True	'Stop current test case execution
		End Select
	End If
	
	Err.Clear
	On Error Goto 0
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to log an error message in the Reported Events sheet in case of error
'Input Parameters 		: blnStopExecution
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 05/01/2015
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_LogError(blnStopExecution)
	RACK_ReportEvent "Error", Err.Description, "Fail"
	
	'BASED ON ERROR NUMBER/DESCRIPTION, ERROR HANDLING CAN BE FURTHER CUSTOMIZED AS FOLLOWS
	'Select Case Err.Description
	'	Case "Object not found"
	'		'Do required error handling
	'	Case "Object disabled"
	'		'Do required error handling
	'	............
	'	............
	'	............
	'End Select
	
	Dim objFso, objMyFile
	Set objFso = CreateObject("Scripting.FileSystemObject")
	If (objFso.FileExists(Pathfinder.Locate("StopAllExecution.txt"))) Then
		Set objMyFile = _
			objFso.OpenTextFile(Pathfinder.Locate("StopAllExecution.txt"), 2)	'Open the StopAllExecution file for writing
		objMyFile.Writeline(blnStopExecution)
		objMyFile.Close
	End If
	
	Set objMyFile = Nothing
	Set objFso = Nothing
End Sub
'#####################################################################################################################

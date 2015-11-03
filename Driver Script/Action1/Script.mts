 'EB'General Header
'#####################################################################################################################
'Script Description		: Driver Script to trigger the testcase execution
'Test Tool/Version		: HP Quick Test Professional 9.5 and above
'Application Automated	: Oracle EBS Application
'Author					: Rackspace
'Date Created			: 12/01/2015
'Created  by : Convene Team
'#####################################################################################################################
'	'Forcing Variable declarations
'SystemUtil.CloseProcessByName("IExplore.exe")
'Declare required variables
'Option Explicit
Dim gobjFso
Dim gstrProjectName
Dim gstrDatatableName
Dim gstrBusinessFlowSheet, gstrCheckPointSheet
Dim gstrResultSheet, gstrReportedEventSheet, gstrCurrentScenario
Dim gstrCurrentTestCase, gstrIterationMode, gintStartIteration
Dim gintEndIteration, gintCurrentIteration, gintCurrentBusinessFlowRow
Dim gintCurrentTestDataRow, gintCurrentReportedEventRow, gintCurrentFlowNumber
Dim garrCurrentFlowData, gstrCurrentKeyword, gintGroupIterations
Dim gintCurrentGroupIteration, garrComponentGroup(), gintGroupedComponents
Dim gintCurrentComponent, gobjHashTable


'Initialize basic configuration settings from RACK.ini file
gstrBusinessFlowSheet =  RACK_GetConfig("BusinessFlowSheet")
gstrCheckPointSheet = RACK_GetConfig("CheckPointSheet")
gstrResultSheet = RACK_GetConfig("ResultSheet")
gstrReportedEventSheet = RACK_GetConfig("ReportedEventSheet")
Environment.Value("CheckpointSheet") = gstrCheckpointSheet
Environment.Value("ReportedEventSheet") = gstrReportedEventSheet
Environment.Value("ResultSheet") = gstrResultSheet
Environment.Value("TakeScreenshotFailedStep") = _
							CBool(RACK_GetConfig("TakeScreenshotFailedStep"))
Environment.Value("TakeScreenshotPassedStep") = _
							CBool(RACK_GetConfig("TakeScreenshotPassedStep"))
'Options are NextIteration, NextTestCase, NextStep, Stop, Dialog
Environment.Value("OnError") = RACK_GetConfig("OnError")	
If CBool(RACK_GetConfig("DebugMode")) Then
	'Turn off error handling to enable debugging
	Environment.Value("OnError") = "NextStep"	
End If
Environment.Value("ReportsTheme") = RACK_GetConfig("ReportsTheme")
Environment.Value("ResultPath") = Pathfinder.Locate("Results")
Environment.Value("OverallStatus") = ""
Environment.Value("RunIndividualComponent") = False
Environment.Value("TestCase_ExecutionTime") = 0

Environment.Value("RequisitionID")=""

'Setup appropriate parameters for the Current Test Case Execution (passed from the initialization script)
gstrCurrentScenario = Parameter("CurrentScenario")
Environment.Value("CurrentScenario") = gstrCurrentScenario
gstrDatatableName = gstrCurrentScenario & ".xls"
gstrCurrentTestCase = Parameter("CurrentTestCase")
Environment.Value("CurrentTestCase") = gstrCurrentTestCase
Environment.Value("TimeStamp") = Parameter("TimeStamp")
gstrIterationMode = Parameter("IterationMode")

'Import required sheets from Datatable
Set gobjFso = CreateObject("Scripting.FileSystemObject")
If Not gobjFso.FileExists(Pathfinder.Locate("Datatables\" & gstrDatatableName)) Then
	Reporter.ReportEvent micFail,"Error",_
						"Datatable not found for the specified Scenario!"
	ExitRun
End If

RACK_ImportSheet Pathfinder.Locate("Datatables\" & gstrDatatableName),_
													gstrBusinessFlowSheet
RACK_ImportSheet Pathfinder.Locate("Datatables\" & gstrDatatableName),_
													gstrReportedEventSheet
If gobjFso.FileExists(Environment.Value("ResultPath") & "\" &_
		Environment.Value("TimeStamp") & "\Excel Results\Summary.xls") Then
	RACK_ImportSheet Environment.Value("ResultPath") & "\" &_
		Environment.Value("TimeStamp") & "\Excel Results\Summary.xls",_
														gstrResultSheet
Else
	RACK_ImportSheet Pathfinder.Locate("Datatables\" & gstrDatatableName),_
																gstrResultSheet
End If
Set gobjFso = Nothing

'Setup the test case iterations
Select Case gstrIterationMode
	Case "oneIteration"
		gintStartIteration = 1
		gintEndIteration = 1
		gintCurrentIteration = 1
	Case "rngIterations"
		gintStartIteration = Parameter("StartIteration")
		gintEndIteration = Parameter("EndIteration")
		If gintStartIteration = "" then
			gintStartIteration = 1
		End if
		If gintEndIteration = "" then
			gintEndIteration = 1
		End if
		gintCurrentIteration = gintStartIteration
	Case "rngAll"
		gintStartIteration = 1
		gintEndIteration = 65535
		gintCurrentIteration = 1
End Select

'Execute all iterations of Current Test Case
Set gobjHashTable = CreateObject("Scripting.Dictionary")
gintCurrentBusinessFlowRow = _
		RACK_SetBusinessFlowRow(gstrCurrentTestCase, gstrBusinessFlowSheet)
Do while gintCurrentIteration <= gintEndIteration
	Environment.Value("CurrentIteration") = gintCurrentIteration
	
	RACK_ReportEvent "Start",_
						"Iteration" & gintCurrentIteration & " started", "Done"
	Environment.Value("Iteration_StartTime") = Now()
	Environment.Value("ExitIteration") = False
	Environment.Value("StopExecution") = False
	
	gintCurrentFlowNumber = 1
	gintGroupedComponents = 0
	garrCurrentFlowData = _
				Split(DataTable.Value("Keyword_1",gstrBusinessFlowSheet),",")
	gstrCurrentKeyword = garrCurrentFlowData(0)
	Do until gstrCurrentKeyword = ""
		If UBound(garrCurrentFlowData) = 0 Then
			gintGroupIterations = 1
		Else
			gintGroupIterations = garrCurrentFlowData(1)
		End If
		
		gintGroupedComponents = gintGroupedComponents + 1
		Redim Preserve garrComponentGroup(gintGroupedComponents)
		garrComponentGroup(gintGroupedComponents - 1) = gstrCurrentKeyword
		
		If (gintGroupIterations > 0) Then	'Reached the end of a group (a group may comprise only one keyword also)
			For gintCurrentGroupIteration = 1 To gintGroupIterations	'Execute all group iterations specified
				For gintCurrentComponent = 0 To (gintGroupedComponents - 1)	'Execute all keywords in the group for the current group iteration
					RACK_ReportEvent "Start Component", "Invoking Business component: " &_
						garrComponentGroup(gintCurrentComponent), "Done"
					
					'Check if the current keyword has already been invoked earlier, and update the hash table accordingly
					If gobjHashTable.Exists(garrComponentGroup(gintCurrentComponent)) Then
						gobjHashTable.Item(garrComponentGroup(gintCurrentComponent)) = _
							gobjHashTable.Item(garrComponentGroup(gintCurrentComponent)) + 1
					Else
						gobjHashTable.Add garrComponentGroup(gintCurrentComponent), 1
					End If

					'Update the current sub iteration number
					Environment.Value("CurrentSubIteration") = gobjHashTable._
								Item(garrComponentGroup(gintCurrentComponent))
					
					RACK_InvokeBusinessComponent garrComponentGroup(gintCurrentComponent)
					
					RACK_ReportEvent "End Component",_
											"Exiting Business component: " &_
							garrComponentGroup(gintCurrentComponent), "Done"
					
					If (Environment.Value("ExitIteration")) Then
						Exit Do
					End If
					
					If (Environment.Value("StopExecution")) Then
						RACK_ReportEvent "RACK_Info", "Execution aborted by user", "Done"
						RACK_CalculateExecTime()
						RACK_WrapUp()
						ExitRun
					End If
				Next
			Next
			
			gintGroupedComponents = 0
		End If
		
		'Process next keyword
		gintCurrentFlowNumber = gintCurrentFlowNumber + 1
		garrCurrentFlowData = Split(DataTable.Value("Keyword_" &_
							gintCurrentFlowNumber,gstrBusinessFlowSheet),",")
		If UBound(garrCurrentFlowData) = -1 Then
			gstrCurrentKeyword = ""
		Else
			gstrCurrentKeyword = garrCurrentFlowData(0)
		End If
	Loop
	
	RACK_ReportEvent "End", "Iteration" &_
									gintCurrentIteration & " completed", "Done"
	RACK_CalculateExecTime()
	
	'Move to the next iteration of test data
	gintCurrentIteration = gintCurrentIteration + 1
	gobjHashTable.RemoveAll()
Loop

Set gobjHashTable = Nothing

RACK_WrapUp()
'############################################################################################################################




'############################################################################################################################
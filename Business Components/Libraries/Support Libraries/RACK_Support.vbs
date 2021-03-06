'General Header
'#####################################################################################################################
'Script Description		: RACK Support Library
'Test Tool/Version		: HP Quick Test Professional 9.5 and above
'Test Tool Settings		: N.A.
'Application Automated	: Flight Application
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Option Explicit	'Forcing Variable declarations

'#####################################################################################################################
'Function Description   : Function to import specified Excel sheet into datatable
'Input Parameters 		: strFilePath, strSheetName
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_ImportSheet (strFilePath, strSheetName)
	Datatable.Addsheet strSheetName
	Datatable.Importsheet strFilePath, strSheetName, strSheetName
End Sub 
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to set the current row in Business Flow Sheet based on the current test case
'Input Parameters 		: strBusinessFlowSheet,strCurrentTestCase
'Return Value    		: intCurrentRow
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Function RACK_SetBusinessFlowRow(strCurrentTestCase, strBusinessFlowSheet)
	Dim intCurrentRow, blnTestCaseFound
	intCurrentRow = 1
	blnTestCaseFound = False
	
	Do until Trim(DataTable.Value("TC_ID",strBusinessFlowSheet)) = ""
		If (DataTable.Value("TC_ID",strBusinessFlowSheet) = strCurrentTestCase) Then
			blnTestCaseFound = True
			Exit Do
		Else
			intCurrentRow = intCurrentRow + 1
			DataTable.GetSheet(strBusinessFlowSheet).SetCurrentRow(intCurrentRow)
		End If
	Loop
	
	If (blnTestCaseFound = False) Then
		RACK_ReportEvent "Error", "Business flow of the test Case " &_
					strCurrentTestCase & " not found in the scenario " &_
							Environment.Value("CurrentScenario"), "Fail"
		'ExitRun
	End If
	
	RACK_SetBusinessFlowRow = intCurrentRow
End Function
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to invoke the corresponding Business component based on the keyword passed
'Input Parameters	 	: strCurrentKeyword
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_InvokeBusinessComponent(strCurrentKeyword)
	If (Environment.Value("OnError") <> "NextStep") Then
		On Error Resume Next
	End If
	
	Execute strCurrentKeyword
	
	RACK_ErrHandler()
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to return the test data value corresponding to the field name passed
'Input Parameters		: strDataSheetName, strFieldName
'Return Value    		: strDataValue
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Function RACK_GetData(strTestDataSheet, strFieldName)
	'Initialise required variables
	Dim strReferenceIdentifier, strCurrentTestCase, intCurrentIteration, intCurrentSubIteration, strDatatableName, strFilePath
	Dim strConnectionString, strSql, objConn, objTestData, strDataValue, strFirstChar
	
	strReferenceIdentifier = RACK_GetConfig("DataReferenceIdentifier")
	strCurrentTestCase = Environment.Value("CurrentTestCase")
	intCurrentIteration = Environment.Value("CurrentIteration")
	intCurrentSubIteration = Environment.Value("CurrentSubIteration")
	strDatatableName = Environment.Value("CurrentScenario") & ".xls"
	strFilePath = PathFinder.Locate("Datatables\" & strDatatableName)
	strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFilePath + ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""
	
	Set objConn = CreateObject("ADODB.Connection")
	objConn.Open strConnectionString
	Set objTestData = CreateObject("ADODB.Recordset")
	objTestData.CursorLocation = 3
	strSql = "SELECT " & strFieldName & " from [" & strTestDataSheet & "$] where TC_ID='" & strCurrentTestCase & "' and Iteration = '" & intCurrentIteration & "' and SubIteration = '" & intCurrentSubIteration & "'"
	objTestData.Open strSql, objConn
    
	If objTestData.RecordCount = 0 Then
		Err.Raise 2001, "RACK", "No test data found for the current row: TC_ID = " &_
				strCurrentTestCase & ", Iteration = " & intCurrentIteration &_
								", SubIteration = " & intCurrentSubIteration
	End If
	strDataValue = Trim(objTestData(0).Value)
	strFirstChar = Left(strDataValue, 1)
	
	If strFirstChar = strReferenceIdentifier Then
		objConn.Close
		strDataValue = Split(strDataValue, strReferenceIdentifier)(1)
		strFilePath = PathFinder.Locate("Datatables\Common Testdata.xls")
		strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFilePath + ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""
		objConn.Open strConnectionString
		strSql = "SELECT " & strFieldName & " from [Common Testdata$] where TD_ID='" & strDataValue & "'"
		objTestData.Open strSql,objConn

	   	If objTestData.RecordCount = 0 Then
			Err.Raise 2002, "RACK", "No common test data found for the current row: TD_ID = " & strDataValue
		End If
		strDataValue = Trim(objTestData(0).Value)
	End If
	
	'Release all objects
	objTestData.Close
	objConn.Close
	Set objConn = Nothing
	Set objTestData = Nothing
	
	'Avoid returning Null value
	If IsNull(strDataValue) Then
		strDataValue = ""
	End If
	RACK_GetData = strDataValue
End Function
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to output intermediate data (output values)  into the Test data sheet
'Input Parameters		: strFieldName, strDataValue
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_PutData(strTestDataSheet, strFieldName, strDataValue)
	'Initialize required variables
	Dim strCurrentTestCase, intCurrentIteration, intCurrentSubIteration, strDatatableName, strFilePath, strConnectionString, objConn, strSql
	If(CBool(Environment.Value("RunIndividualComponent")) <> True) Then
		strCurrentTestCase = Environment.Value("CurrentTestCase")
		intCurrentIteration = Environment.Value("CurrentIteration")
		intCurrentSubIteration = Environment.Value("CurrentSubIteration")
		strDatatableName = Environment.Value("CurrentScenario") & ".xls"
		strFilePath = PathFinder.Locate("Datatables\" & strDatatableName)
		strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFilePath + ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""
		
		'Write the output value into the test data sheet
		Set objConn = CreateObject("ADODB.Connection")
		objConn.Open strConnectionString
		strSql = "UPDATE [" & strTestDataSheet & "$] SET " & strFieldName & "='" & strDataValue & "' where TC_ID='" & strCurrentTestCase & "' and Iteration = '" & intCurrentIteration & "' and SubIteration = '" & intCurrentSubIteration & "'"
		objConn.Execute strSql
		objConn.Close
		Set objConn = Nothing
		
		'Report the output value to the results	
		RACK_ReportEvent "Output Value", "Output value '" & strDataValue & "' written into the '" & strFieldName & "' column", "Done"
	End If
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to return the expected result data (from the Parameterized Checkpoints sheet) corresponding to the field name passed
'Input Parameters		: strFieldName
'Return Value    		: strDataValue
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Function RACK_GetExpectedResult(strFieldName)
	'Initialise required variables
	Dim strReferenceIdentifier, strCheckPointSheet, strCurrentTestCase, intCurrentIteration, strDatatableName, strFilePath
	Dim strConnectionString, strSql, objConn, objTestData, strDataValue, strFirstChar
	
	strReferenceIdentifier = RACK_GetConfig("DataReferenceIdentifier")
	strCheckPointSheet = Environment.Value("CheckpointSheet")
	strCurrentTestCase = Environment.Value("CurrentTestCase")
	intCurrentIteration = Environment.Value("CurrentIteration")
	intCurrentSubIteration = Environment.Value("CurrentSubIteration")
	strDatatableName = Environment.Value("CurrentScenario") & ".xls"
	strFilePath = PathFinder.Locate("Datatables\" & strDatatableName)
	strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFilePath + ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""
	
	Set objConn = CreateObject("ADODB.Connection")
	objConn.Open strConnectionString
	Set objTestData = CreateObject("ADODB.Recordset")
	objTestData.CursorLocation = 3	
	strSql = "SELECT " & strFieldName & " from [" & strCheckPointSheet & "$] where TC_ID='" & strCurrentTestCase & "' and Iteration = '" & intCurrentIteration & "' and SubIteration = '" & intCurrentSubIteration & "'"
	objTestData.Open strSql, objConn
	
	If objTestData.RecordCount = 0 Then
		Err.Raise 2003, "RACK", "No expected results found for the current row: TC_ID = " & strCurrentTestCase & ", Iteration = " & intCurrentIteration & ", SubIteration = " & intCurrentSubIteration
	End If
	strDataValue = Trim(objTestData(0).Value)
	
	'Release all objects
	objTestData.Close
	objConn.Close
	Set objConn = Nothing
	Set objTestData = Nothing
	
	'Avoid returning Null value
	If IsNull(strDataValue) Then
		strDataValue = ""
	End If
	RACK_GetExpectedResult = strDataValue
End Function
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to get the configuration data from the RACK.ini configuration file
'Input Parameters		: strKey
'Return Value    		: Corresponding value from RACK.ini
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Function RACK_GetConfig(strKey)
	Dim objFso, objMyFile
	Dim strLine, arrLine, strValue, strConfigFilePath
	Set objFso = CreateObject ("Scripting.FileSystemObject")
	strConfigFilePath = PathFinder.Locate("RACK.ini")
	Set objMyFile = objFso.OpenTextFile(strConfigFilePath,1)
	Do Until objMyFile.AtEndOfStream
		strLine = objMyFile.ReadLine
		If strLine <> "" Then
			arrLine = Split(strLine,"=")
			If arrLine(0) = strKey Then
				strValue = arrLine(1)
				Exit Do
			End If
		End If
	Loop
	
	objMyFile.Close()
	Set objMyFile = Nothing
	Set objFso = Nothing
	RACK_GetConfig = CStr(strValue)
End Function
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to do calculate the execution time for the current iteration
'Input Parameters 		: None
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_CalculateExecTime()
	Dim dtmIteration_EndTime, dtmIteration_StartTime, sngIteration_ExecutionTime
	Dim strReportedEventSheet,intCurrentReportedEventRow
	dtmIteration_StartTime = Environment.Value("Iteration_StartTime")
	dtmIteration_EndTime = Now()
	strReportedEventSheet = Environment.Value("ReportedEventSheet")
	
	'Report the total execution time for the current iteration and insert a blank row
	sngIteration_ExecutionTime = DateDiff("s", dtmIteration_StartTime, dtmIteration_EndTime)
	sngIteration_ExecutionTime = Round(CSng(sngIteration_ExecutionTime)/60, 2)
	DataTable.Value("Description", strReportedEventSheet) = "Execution Time (mins)"
	DataTable.Value("Time", strReportedEventSheet) = sngIteration_ExecutionTime
	Environment.Value("TestCase_ExecutionTime") = _
		Environment.Value("TestCase_ExecutionTime") + sngIteration_ExecutionTime
	intCurrentReportedEventRow = _
					DataTable.GetSheet(strReportedEventSheet).GetCurrentRow()
	DataTable.GetSheet(strReportedEventSheet).SetCurrentRow(intCurrentReportedEventRow + 2)
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to do required wrap-up work after running a test case
'Input Parameters 		: strDatatableName
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_WrapUp()
	'Initialise required variables
	Dim strProjectName, strCurrentTestCase, strDescription, strReportsTheme
	strProjectName = RACK_GetConfig("ProjectName")
	strCurrentTestCase = Parameter("CurrentTestCase")
	strDescription = Parameter("TestCaseDescription")
	strReportsTheme = Environment.Value("ReportsTheme")
	
	'Update overall result of the test case
	If (Environment.Value("OverallStatus") <> "Fail") Then
		Environment.Value("OverallStatus") = "Pass"
	End If
	
	'Export Results to Excel and HTML
	RACK_ExportReportedEventsToExcel strCurrentTestCase
	RACK_ExportReportedEventsToHtml strProjectName, strCurrentTestCase, strReportsTheme
	RACK_UpdateResultSummary strCurrentTestCase, strDescription, Environment.Value("TestCase_ExecutionTime")
	RACK_ExportResultSummaryToExcel()
	RACK_ExportResultSummaryToHtml strProjectName, strReportsTheme
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to report any event related to the current test case
'Input Parameters 		: strStepName, strDescription, strStatus
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_ReportEvent(strStepName, strDescription, strStatus)
	'Report the event in QTP results
	Dim intStatus
	Select Case strStatus
		Case "Pass"
			intStatus = 0
		Case "Fail"
			intStatus = 1
		Case "Done"
			intStatus = 2
		Case "Warning"
			intStatus = 3
		Case "Screenshot"
			intStatus = 4	
	End Select
	Reporter.ReportEvent intStatus,strStepName,strDescription
	
	'Report the event in Excel/HTML results
	If(CBool(Environment.Value("RunIndividualComponent")) <> True) Then
		Dim strReportedEventSheet, strCurrentTime
		strReportedEventSheet = Environment.Value("ReportedEventSheet")
		strCurrentTime = Time()
		DataTable.Value("Iteration",strReportedEventSheet) = Environment.Value("CurrentIteration")
		DataTable.Value("Step_Name",strReportedEventSheet) = strStepName
		DataTable.Value("Description",strReportedEventSheet) = strDescription
		DataTable.Value("Status",strReportedEventSheet) = strStatus
		DataTable.Value("Time",strReportedEventSheet) = strCurrentTime
		
		Dim objFso, strScreenshotPath
		Set objFso = CreateObject("Scripting.FileSystemObject")
		strScreenshotPath = Environment.Value("ResultPath") & "\" &_
							Environment.Value("TimeStamp") & "\Screenshots\" &_
								Parameter("CurrentTestCase") & "_Iteration" &_
								Environment.Value("CurrentIteration") & "_" &_
								Replace(strCurrentTime,":","-") &".png"
		
		'Take screenshot if its a failed step or a warning (only if the user has enabled this setting, and another screenshot was not taken already in the very same second)
		If((strStatus = "Fail" Or strStatus = "Warning") And Environment.Value("TakeScreenshotFailedStep")) And objFso.FileExists(strScreenshotPath) = False Then
			Desktop.CaptureBitmap(strScreenshotPath)
		End If
		
		'Take screenshot if its a passed step (only if the user has enabled this setting, and another screenshot was not taken already in the very same second)
		If((strStatus = "Pass") And Environment.Value("TakeScreenshotPassedStep")) And objFso.FileExists(strScreenshotPath) = False Then
			Desktop.CaptureBitmap(strScreenshotPath)
		End If
		
		'Take screenshot if the user requires this step and another screenshot was not taken already in the very same second
		If(strStatus = "Screenshot") And objFso.FileExists(strScreenshotPath) = False Then
			Desktop.CaptureBitmap(strScreenshotPath)
		End If
		
		Set objFso = Nothing
		
		'Set next row in the Reported Events sheet
		Dim intCurrentRow
		intCurrentRow = DataTable.GetSheet(strReportedEventSheet).GetCurrentRow()
		DataTable.GetSheet(strReportedEventSheet).SetCurrentRow(intCurrentRow + 1)
		
		'Update the overall status of the test case
		If(Environment.Value("OverallStatus") <> "Fail") Then
			If(strStatus="Fail") Then
				Environment.Value("OverallStatus") = "Fail"
			End If
		End If
	End If
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to export the reported events in the test case to Excel
'Input Parameters 		: strCurrentTestCase, strReportedEventSheet
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_ExportReportedEventsToExcel(strCurrentTestCase)
	DataTable.ExportSheet Environment.Value("ResultPath") & "\" & Environment.Value("TimeStamp") & "\Excel Results\" & strCurrentTestCase & ".xls", Environment.Value("ReportedEventSheet")
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to set the colors of the HTML report based on the theme specified by the user
'Input Parameters	 	: strReportsTheme, strHeadingColor, strSettingColor, strBodyColor 
'Return Value    		: strHeadingColor, strSettingColor, strBodyColor (through ByRef)
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team                     
'#####################################################################################################################
Sub RACK_SetReportsTheme(strReportsTheme, ByRef strHeadingColor, ByRef strSettingColor, ByRef strBodyColor)
    'Themes can be easily extended by expanding this function
    Select Case UCase(strReportsTheme)
        Case "AUTUMN"
            strHeadingColor="#4bacc6"
            strSettingColor="#d0e3ea"
            strBodyColor="#e9f1f5"
        Case "OLIVE"
            strHeadingColor="#8064a2"
            strSettingColor="#edeaf0"
            strBodyColor="#d8d3e0"
		Case "CLASSIC"
			strHeadingColor="#687C7D"
			strSettingColor="#C6D0D1"
			strBodyColor="#EDEEF0"
        Case "RETRO"
            strHeadingColor="#4f81bd"
            strSettingColor="#e9edf4"
            strBodyColor="#d0d8e8"
        Case "MYSTIC"
            strHeadingColor="#000000"
            strSettingColor="#e7e7e7"
            strBodyColor="#cbcbcb"    
        Case "SERENE"
            strHeadingColor="#000000"
            strSettingColor="#4f81bd"
            strBodyColor="#3d6696"
        Case "REBEL"
            strHeadingColor="#c0504d"
            strSettingColor="#ffffff"
            strBodyColor="#f4e9e9"
        Case Else
            strHeadingColor="#9bbb59"
            strSettingColor="#eff3ea"
            strBodyColor="#dee7d1"    
    End Select
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to export the reported events in the test case to Html
'Input Parameters		: strProjectName, strCurrentTestCase, strReportedEventSheet, strReportsTheme
'Return Value           : None
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_ExportReportedEventsToHtml(strProjectName, strCurrentTestCase, strReportsTheme)
	Dim objFso, objMyFile
	Dim strReportedEventSheet
	Dim intPassCounter, intFailCounter, intVerificationNo
	Dim strIteration, strStepName, strDescription
	Dim strStatus, strTime, strExecutionTime
	Dim intRowcount, intRowCounter, strTempStatus
	Dim strPath, strScreenShotPath, strScreenShotName
	Dim arrSplitTimeStamp, strTimeStampDate, strTimeStampTime
	Dim strOnError, strIterationMode, intStartIteration, intEndIteration
	Dim strHeadColor, strSettColor, strContentBGColor
	
	strReportedEventSheet = Environment.Value("ReportedEventSheet")
	arrSplitTimeStamp = Split(Environment.Value("TimeStamp"),"_")
	strTimeStampDate = Replace(arrSplitTimeStamp(1),"-","/")
	strTimeStampTime = Replace(arrSplitTimeStamp(2),"-",":")
	
	strPath = Environment.Value("ResultPath") & "\" &_
				Environment.Value("TimeStamp") & "\HTML Results\" &_
										strCurrentTestCase & ".html"	

	strScreenShotPath = "..\Screenshots\"
	intPassCounter = 0
	intFailCounter = 0
	intVerificationNo = 0
	
	strOnError = Environment.Value("OnError")
	
	Select Case Parameter("IterationMode")
		Case "oneIteration"
			strIterationMode = "Run one iteration only"
		Case "rngIterations"
			strIterationMode = "Run from <i>Start Iteration</i> to <i>End Iteration</i>"
			
			intStartIteration = Parameter("StartIteration")
			intEndIteration = Parameter("EndIteration")
			If intStartIteration = "" Then
				intStartIteration = 1
			End if
			If intEndIteration = "" Then
				intEndIteration = 1
			End if
		Case "rngAll"
			strIterationMode = "Run all iterations"
	End Select
            strHeadingColor="#4f81bd"
            strSettingColor="#e9edf4"
            strBodyColor="#d0d8e8"
	
	RACK_SetReportsTheme strReportsTheme,strHeadColor,strSettColor,strContentBGColor
	
	'Create a HTML file
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set objMyFile = objFso.CreateTextFile(strPath,True)
	objMyFile.Close

	'Open the HTML file for writing
	Set objMyFile = objFso.OpenTextFile(strPath,8)

	'Create the Report header

    	objMyFile.Writeline("<html>")
		objMyFile.Writeline("<head>")
		objMyFile.Writeline("<meta http-equiv=" & "Content-Language" & "content=" & "en-us>")
		objMyFile.Writeline("<meta http-equiv="& "Content-Type" & "content=" & "text/html; charset=windows-1252" & ">")
		objMyFile.Writeline("<title> Automation Execution Results</title>")
		objMyFile.Writeline("</head>")
		
		objMyFile.Writeline("<body bgcolor = #FFFFFF>")
			objMyFile.Writeline("<blockquote>")
                objMyFile.Writeline("<p align = center><table border=1 bordercolor=" & "#000000 id=table1 width=800 height=31 cellspacing=0 bordercolorlight=" & "#FFFFFF>")
				
				objMyFile.Writeline("<tr height=30px >")
					objMyFile.Writeline("<td COLSPAN = 6 bgcolor =" & strHeadColor &">")
						objMyFile.Writeline("<p align=center><font color=#000000 size=4 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">&nbsp; Automation Execution Results - " & strProjectName  & "</font><font face= " & chr(34)&"Copperplate Gothic"&chr(34) & "></font> </p>")
					objMyFile.Writeline("</td>")
				objMyFile.Writeline("</tr>")
				
				objMyFile.Writeline("<tr>")
					objMyFile.Writeline("<td COLSPAN = 2 bgcolor = " & strSettColor &">")
						'objMyFile.Writeline("<p align=left><font color=#0000000  size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Date: " & strTimeStampDate & " " & strTimeStampTime & "</font><font face= " & chr(34)&"Copperplate Gothic"&chr(34) & "></font> </p>")
						objMyFile.Writeline("<p align=LEFT><font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">"& "&nbsp;"& "DATE: " &  strTimeStampDate & " " )	
					objMyFile.Writeline("</td>")
					
					objMyFile.Writeline("<td COLSPAN = 3 bgcolor = " & strSettColor &">")
						'objMyFile.Writeline("<p align=right><font color=#0000000  size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Total Execution Time: " & intTotalExecTime & " " & strUnit  & "</font> </p>")
						objMyFile.Writeline("<p align=RIGHT><font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">"& "&nbsp;"& "TIME: " & strTimeStampTime)
					objMyFile.Writeline("</td>")
				objMyFile.Writeline("</tr>")
				
				objMyFile.Writeline("<tr height=30px   size=2 bgcolor=" & strHeadColor &">"   )
				objMyFile.Writeline("<td width=" & "400")
						objMyFile.Writeline("<p align=" & "center><font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Process Flow")
				objMyFile.Writeline("</td>")

					objMyFile.Writeline("<td width=" & "400")
						objMyFile.Writeline("<p align=" & "center><font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Test Case ID")
					objMyFile.Writeline("</td>")
					
					objMyFile.Writeline("<td width=" & "400")
						objMyFile.Writeline("<p align=" & "center> <font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Test case Name")
					objMyFile.Writeline("</td>")
					
					objMyFile.Writeline("<td width=" & "400")
						objMyFile.Writeline("<p align=" & "center> <font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Execution Time (Minutes)")
					objMyFile.Writeline("</td>")
					
					objMyFile.Writeline("<td width=" & "400")
						objMyFile.Writeline("<p align=" & "center> <font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & "> " & "Status")
					objMyFile.Writeline("</td>")
				objMyFile.Writeline("</tr>")
					
'					objMyFile.Writeline("<table border=0 bordercolor=" & "#000000 id=table1 width=800 height=31 cellspacing=0 bordercolorlight=" & "#FFFFFF>")				
'					objMyFile.Writeline("<tr bgcolor = "& strSettColor & ">")
'						objMyFile.Writeline("<td colspan =2>")
'							objMyFile.Writeline("<p align=justify><b><font color=" & strHeadColor & " size=2 face= Verdana>"& "&nbsp;"& "OnError: " & strOnError)
'						objMyFile.Writeline("</td>")					  
'						
'						objMyFile.Writeline("<td colspan =2>")
'							objMyFile.Writeline("<p align=right><b><font color=" & strHeadColor & " size=2 face= Verdana>"& "&nbsp;"& "IterationMode: " &  strIterationMode )
'						objMyFile.Writeline("</td>") 
'					objMyFile.Writeline("</tr>") 	   
					
'					If Parameter("IterationMode")="rngIterations" Then
'						objMyFile.Writeline("<tr bgcolor = "& strSettColor & ">")
'							objMyFile.Writeline("<td COLSPAN = 4>")
'								Dim strESpace, strSSpace
'								strESpace=" "
'								strSSpace=";"
'								objMyFile.Writeline("<p align=justify><b><font color=" & strHeadColor & " size=2 face= Verdana>"& strSSpace & "Start Iteration: " & intStartIteration & strESpace & "End Iteration: " & intEndIteration)
'							objMyFile.Writeline("</td>")					  
'						objMyFile.Writeline("</tr>") 	   
'					End If

''Summary 
'Add the data from the Summary file to the HTML file
				Dim strResultSheet

 Dim SheetVal

	Dim objExcel, objWorkBook, objWorkSheet

	strResultSheet = Environment.Value("ResultSheet")

   Set objExcel = CreateObject("Excel.Application")

 Set objWorkBook = objExcel.Workbooks.Open(Pathfinder.Locate("Datatables\RackSpace_Macro.xlsm"))

	
 Set objWorkSheet = objWorkBook.WorkSheets("Sheet1") 
  objWorkSheet.Activate


 SheetVal =objWorkSheet.Range("A" & "1").Value

   objWorkBook.Save
objWorkBook.Close
	
	'objExcel.quit
	
	'Release all objects
	Set objWorkSheet = Nothing
	Set objWorkBook = Nothing
	Set objExcel = Nothing


				strResultSheet = Environment.Value("ResultSheet")
				intRowCount = Datatable.GetSheet(strResultSheet).GetRowCount
				For intRowCounter = 1 To intRowCount
					Datatable.GetSheet(strResultSheet).SetCurrentRow(intRowCounter)	
					strTC_ID = Datatable("TC_ID", strResultSheet)
					strDescription = Datatable("Description", strResultSheet)
					strExecutionTime = Datatable("Execution_Time_Minutes", strResultSheet)
					strStatus = Datatable("Status", strResultSheet)
					strLnkFileName = strTC_ID

					objMyFile.Writeline("<tr bgcolor = " & strSettColor & ">")	
					   objMyFile.Writeline("<td width=" & "400>")							
				        	objMyFile.Writeline("<p align=center> <font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">"  & SheetVal)
						objMyFile.Writeline("</td>")
						objMyFile.Writeline("<td width=" & "400>")							
				        	objMyFile.Writeline("<p align=center> <font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">"  & strTC_ID)
						objMyFile.Writeline("</td>")
						
						objMyFile.Writeline("<td width=" & "400>")
					        objMyFile.Writeline("<p align=" & "center> <font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">"  &  strDescription)
						objMyFile.Writeline("</td>")			
						
						objMyFile.Writeline("<td width=" & "400>")
					        objMyFile.Writeline("<p align=" & "center> <font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">"  &  strExecutionTime)
						objMyFile.Writeline("</td>")		
						
						objMyFile.Writeline("<td width=" & "400>")
							If UCase(strStatus) = "PASS" Then
								objMyFile.Writeline("<p align=" & "center" & ">" & "<font face=" & "Verdana " & "size=" & "2" & " color=" & "#008000" & ">" & strStatus & "</font>")
								intPassCounter = intPassCounter + 1
							ElseIf UCase(strStatus) = "FAIL" Then
								objMyFile.Writeline("<p align=" & "center" & ">" & "<font face=" & "Verdana " & "size=" & "2" & " color=" & "#FF0000" & ">" & strStatus & "</font>")
								intFailCounter = intFailCounter + 1
							Else
								objMyFile.Writeline("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#8A4117" & ">" & strStatus & "</font>")
								intNoRunCounter=intNoRunCounter + 1
							End If
						objMyFile.Writeline("</td>")			
					objMyFile.Writeline("</tr>")	
				Next
''Summary End

			''	objMyFile.Writeline("<table>")				


  			'' objMyFile.Writeline("<p align = center><table border=1 bordercolor=" & "#000000 id=table1 width=800 height=31 cellspacing=0 bordercolorlight=" & "#FFFFFF>")
					objMyFile.Writeline("<tr bgcolor=" & strHeadColor & ">")

						objMyFile.Writeline("<td width=" & "400")
							objMyFile.Writeline("<p align=" & "center> <font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Iteration")
						objMyFile.Writeline("</td>")
'						
						objMyFile.Writeline("<td width=" & "400")
							objMyFile.Writeline("<p align=" & "center> <font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Step Name")
						objMyFile.Writeline("</td>")
						
						objMyFile.Writeline("<td width=" & "400")
							objMyFile.Writeline("<p align=" & "center> <font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Description")
						objMyFile.Writeline("</td>")
						
						objMyFile.Writeline("<td width=" & "400")
							objMyFile.Writeline("<p align=" & "center> <font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Results")
						objMyFile.Writeline("</td>")
'						
						objMyFile.Writeline("<td width=" & "400")
							objMyFile.Writeline("<p align=" & "center> <font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Time")
						objMyFile.Writeline("</td>")
					objMyFile.Writeline("</tr>")
				'End of Header
				
				'Add Data to the Test Case Log HTML file from the excel file
					intRowcount = Datatable.GetSheet(strReportedEventSheet).GetRowCount
					For intRowCounter = 1 To intRowCount
						Datatable.GetSheet(strReportedEventSheet).SetCurrentRow(intRowCounter)	
						strIteration = Datatable("Iteration",strReportedEventSheet)
						strStepName = Datatable("Step_Name",strReportedEventSheet)
						strDescription = Datatable("Description",strReportedEventSheet)
						strStatus = Datatable("Status",strReportedEventSheet)
						strTime = Datatable("Time",strReportedEventSheet)
						
						If strIteration = "" Then
							objMyFile.Writeline("<tr bgcolor =" & strContentBGColor & ">")
								objMyFile.Writeline("<td COLSPAN = 6>")
									objMyFile.Writeline("<p align=center><b><font size=2 face= Verdana>"& "&nbsp;"& strDescription & ":&nbsp;&nbsp;" &  strTime  & "&nbsp")
								objMyFile.Writeline("</td>")
							objMyFile.Writeline("</tr>")
							intRowCounter = intRowCounter+1
						Else
							objMyFile.Writeline("<tr bgcolor =" & strContentBGColor & ">")
							objMyFile.Writeline("<td width=" & "400>")
									objMyFile.Writeline("<p align=" & "center><font face=" & "Verdana " & "size=" & "2" & ">"  &  strIteration)
								objMyFile.Writeline("</td>")
								
								objMyFile.Writeline("<td width=" & "400>")
									strScreenShotName = Parameter("CurrentTestCase") & "_Iteration" & strIteration & "_" & Replace(strTime,":","-")
									If(UCase(strStatus) = "FAIL" And Environment.Value("TakeScreenshotFailedStep")) Then										
										objMyFile.Writeline("<p align=center><a href='" & strScreenShotPath & strScreenShotName & ".png" & "'><b><font face=" & "verdana" & "size=" & "2" & ">" & strStepName & "</font></b></a></p>")
									ElseIf(UCase(strStatus) = "PASS" And Environment.Value("TakeScreenshotPassedStep")) Then										
										objMyFile.Writeline("<p align=center><a href='" & strScreenShotPath & strScreenShotName & ".png" & "'><b><font face=" & "verdana" & "size=" & "2" & ">" & strStepName & "</font></b></a></p>")
									ElseIf(UCase(strStatus) = "SCREENSHOT") Then
										objMyFile.Writeline("<p align=center><a href='" & strScreenShotPath & strScreenShotName & ".png" & "'><b><font face=" & "verdana" & "size=" & "2" & ">" & strStepName & "</font></b></a></p>")
									Else
										objMyFile.Writeline("<p align=" & "center><font face=" & "Verdana " & "size=" & "2" & ">"  &  strStepName)
									End If
								objMyFile.Writeline("</td>")
								
								objMyFile.Writeline("<td width=" & "400>")
									objMyFile.Writeline("<p align=" & "center><font face=" & "Verdana " & "size=" & "2" & ">"  &  strDescription)
								objMyFile.Writeline("</td>")
								
								objMyFile.Writeline("<td width=" & "400>")
									If UCase(strStatus) = "PASS" Then
										objMyFile.Writeline("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#008000" & ">" & strStatus & "</font></b>")
										intPassCounter=intPassCounter + 1	
										intVerificationNo=intVerificationNo + 1
									ElseIf UCase(strStatus) = "FAIL" Then
										objMyFile.Writeline("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#FF0000" & ">" & strStatus & "</font></b>")
										intFailCounter=intFailCounter + 1
										intVerificationNo=intVerificationNo + 1
									ElseIf UCase(strStatus) = "SCREENSHOT" Then
										objMyFile.Writeline("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#003399" & ">" & strStatus & "</font></b>")									
									Else
										objMyFile.Writeline("<p align=" & "center" & ">" & "<b><font face=" & "Verdana " & "size=" & "2" & " color=" & "#8A4117" & ">" & strStatus & "</font></b>")		
									End If
								objMyFile.Writeline("</td>")
							
								objMyFile.Writeline("<td width=" & "400>")
									objMyFile.Writeline("<p align=" & "center><font face=" & "Verdana " & "size=" & "2" & ">"  &  strTime)
								objMyFile.Writeline("</td>")
							objMyFile.Writeline("</tr>")
						End If
					Next
			'	objMyFile.Writeline("</table>")				
				
			'	objMyFile.Writeline("<table border=1 bordercolor=" & "#000000 id=table1 width=800 height=31 cellspacing=0 bordercolorlight=" & "#FFFFFF>")
					objMyFile.Writeline("<tr bgcolor =" & strHeadColor & ">")
'						objMyFile.Writeline("<td colspan =1>")
'							objMyFile.Writeline("<p align=justify><b><font color=#000000  size=2 face= Verdana>"& "&nbsp;"& "No. Of Verification Points :&nbsp;&nbsp;" &  intVerificationNo & "&nbsp;")
'						objMyFile.Writeline("</td>")
						
						objMyFile.Writeline("<td colspan =1>")	
							objMyFile.Writeline("<p align=LEFT><b><font color=#000000  size=2 face= Verdana>"& "&nbsp;"& "Passed :&nbsp;&nbsp;" &  intPassCounter & "&nbsp;")
						objMyFile.Writeline("</td>")

						objMyFile.Writeline("<td colspan =1>")
						objMyFile.Writeline("<p align=justify><b><font color=#000000  size=2 face= Verdana>"& "&nbsp;"& "")
						objMyFile.Writeline("</td>")

						objMyFile.Writeline("<td colspan =1>")
						objMyFile.Writeline("<p align=justify><b><font color=#000000  size=2 face= Verdana>"& "&nbsp;"& "")
						objMyFile.Writeline("</td>")

						objMyFile.Writeline("<td colspan =1>")
						objMyFile.Writeline("<p align=justify><b><font color=#000000  size=2 face= Verdana>"& "&nbsp;"& "")
						objMyFile.Writeline("</td>")
						
						objMyFile.Writeline("<td colspan =1>")	
							objMyFile.Writeline("<p align=RIGHT><b><font color=#000000  size=2 face= Verdana>"& "&nbsp;"& "Failed :&nbsp;&nbsp;" &  intFailCounter & "&nbsp;")
						objMyFile.Writeline("</td>")	
					objMyFile.Writeline("</tr>")	
				objMyFile.Writeline("</table>")				
			objMyFile.Writeline("</blockquote>")			
		objMyFile.Writeline("</body>")
	objMyFile.Writeline("</html>")
	objMyFile.Close
	
	Set objMyFile = Nothing
	Set objFso = Nothing
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to update the Results Summary with the current Test Case Iteration status
'Input Parameters	 	: strCurrentTestCase, strDescription, sngExecutionTime, strResultSheet
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_UpdateResultSummary(strCurrentTestCase, strDescription, sngExecutionTime)
	Dim strResultSheet
	strResultSheet = Environment.Value("ResultSheet")
	DataTable.GetSheet(strResultSheet).SetCurrentRow(DataTable.GetSheet(strResultSheet).GetRowCount+1)
	DataTable.Value("TC_ID",strResultSheet) = strCurrentTestCase
	DataTable.Value("Description",strResultSheet) = strDescription
	DataTable.Value("Execution_Time_Minutes",strResultSheet) = sngExecutionTime
	DataTable.Value("Status",strResultSheet) = Environment.Value("OverallStatus")
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to exported the Results Summary sheet to Excel
'Input Parameters 		: strResultSheet
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_ExportResultSummaryToExcel()
	DataTable.ExportSheet Environment.Value("ResultPath") & "\" &_
					Environment.Value("TimeStamp") &_
					"\Excel Results\Summary.xls",_
					Environment.Value("ResultSheet")
End Sub
'#####################################################################################################################

'#####################################################################################################################
'Function Description   : Function to exported the Results Summary sheet to HTML
'Input Parameters 		: strResultSheet
'Return Value    		: None
'Author					: Rackspace
'Date Created			: 27/12/2014
'Created  by : Convene Team
'#####################################################################################################################
Sub RACK_ExportResultSummaryToHtml(strProjectName, strReportsTheme)
	Dim objFso, objMyFile
	Dim strResultSheet
	Dim intPassCounter, intFailCounter, intNoRunCounter
	Dim intRowCount, intRowCounter
	Dim strTC_ID, strDescription, strExecutionTime, strStatus
	Dim strLnkFileName, strPath
	Dim intTotalExecTime, strExecTimeTemp, strUnit
	Dim arrSplitTimeStamp, strTimeStampDate, strTimeStampTime
	Dim strHeadColor, strSettColor, strContentBGColor
    Dim SheetVal

	Dim objExcel, objWorkBook, objWorkSheet

	strResultSheet = Environment.Value("ResultSheet")

   Set objExcel = CreateObject("Excel.Application")

 Set objWorkBook = objExcel.Workbooks.Open(Pathfinder.Locate("Datatables\RackSpace_Macro.xlsm"))

	
 Set objWorkSheet = objWorkBook.WorkSheets("Sheet1") 
  objWorkSheet.Activate


 SheetVal =objWorkSheet.Range("A" & "1").Value

   objWorkBook.Save
objWorkBook.Close
	
	'objExcel.quit
	
	'Release all objects
	Set objWorkSheet = Nothing
	Set objWorkBook = Nothing
	Set objExcel = Nothing



	arrSplitTimeStamp = Split(Environment.Value("TimeStamp"),"_")
	strTimeStampDate = Replace(arrSplitTimeStamp(1),"-","/")
	strTimeStampTime = Replace(arrSplitTimeStamp(2),"-",":")	
	intPassCounter = 0
	intFailCounter = 0
	intNoRunCounter = 0
	intTotalExecTime = 0
	strPath = Environment.Value("ResultPath") & "\" &_
				Environment.Value("TimeStamp") & "\HTML Results\Summary.html"
	
	'Default settings for theme
            strHeadingColor="#4f81bd"
            strSettingColor="#e9edf4"
            strBodyColor="#d0d8e8"
	
	RACK_SetReportsTheme strReportsTheme, strHeadColor, strSettColor, strContentBGColor
	
	'Count the total Execution time
	intRowCount = Datatable.GetSheet(strResultSheet).GetRowCount
	For intRowCounter = 1 To intRowCount
		Datatable.GetSheet(strResultSheet).SetCurrentRow(intRowCounter)		
		strExecTimeTemp = Datatable("Execution_Time_Minutes",strResultSheet)
		intTotalExecTime = intTotalExecTime+CSng(strExecTimeTemp)	
	Next
	
	If intTotalExecTime = 1 Then
		strUnit = "minute"
	Else
		strUnit = "minutes"
	End If
	
	'Create a HTML file
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set objMyFile = objFso.CreateTextFile(strPath, True)
	objMyFile.Close
	
	'Open the HTML file for writing
	Set objMyFile = objFso.OpenTextFile(strPath,8)

	'Create the Report header
	objMyFile.Writeline("<html>")
		objMyFile.Writeline("<head>")
			objMyFile.Writeline("<meta http-equiv=" & "Content-Language" & "content=" & "en-us>")
			objMyFile.Writeline("<meta http-equiv="& "Content-Type" & "content=" & "text/html; charset=windows-1252" & ">")
			objMyFile.Writeline("<title> Automation Execution Results</title>")
		objMyFile.Writeline("</head>")
		
		objMyFile.Writeline("<body bgcolor = #FFFFFF>")
			objMyFile.Writeline("<blockquote>")
                objMyFile.Writeline("<p align = center><table border=1 bordercolor=" & "#FFFFFF id=table1 width=800 height=31 cellspacing=0 bordercolorlight=" & "#FFFFFF>")
				
				objMyFile.Writeline("<tr height=30px >")
					objMyFile.Writeline("<td COLSPAN = 6 bgcolor =" & strHeadColor &">")
						objMyFile.Writeline("<p align=center><font color=#000000 size=4 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">&nbsp; Automation Execution Results - " & strProjectName  & "</font><font face= " & chr(34)&"Copperplate Gothic"&chr(34) & "></font> </p>")
					objMyFile.Writeline("</td>")
				objMyFile.Writeline("</tr>")
				
				objMyFile.Writeline("<tr>")
					objMyFile.Writeline("<td COLSPAN = 2 bgcolor = " & strSettColor &">")
						objMyFile.Writeline("<p align=left><font color=#0000000  size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Date: " & strTimeStampDate & " " & strTimeStampTime & "</font><font face= " & chr(34)&"Copperplate Gothic"&chr(34) & "></font> </p>")
					objMyFile.Writeline("</td>")
					
					objMyFile.Writeline("<td COLSPAN =3 bgcolor = " & strSettColor &">")
						objMyFile.Writeline("<p align=right><font color=#0000000  size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Total Execution Time: " & intTotalExecTime & " " & strUnit  & "</font> </p>")
					objMyFile.Writeline("</td>")
				objMyFile.Writeline("</tr>")


				
				objMyFile.Writeline("<tr height=30px   size=2 bgcolor=" & strHeadColor &">"   )

					objMyFile.Writeline("<td width=" & "400")
			objMyFile.Writeline("<p align=" & "center><font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Process Flow")
			objMyFile.Writeline("</td>")

					objMyFile.Writeline("<td width=" & "400")
						objMyFile.Writeline("<p align=" & "center><font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Test Case ID")
					objMyFile.Writeline("</td>")
					
					objMyFile.Writeline("<td width=" & "400")
						objMyFile.Writeline("<p align=" & "center> <font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Test case Name")
					objMyFile.Writeline("</td>")
					
					objMyFile.Writeline("<td width=" & "400")
						objMyFile.Writeline("<p align=" & "center> <font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">" & "Execution Time (Minutes)")
					objMyFile.Writeline("</td>")
					
					objMyFile.Writeline("<td width=" & "400")
						objMyFile.Writeline("<p align=" & "center> <font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & "> " & "Status")
					objMyFile.Writeline("</td>")
				objMyFile.Writeline("</tr>")
				'End of Header
				
				'Add the data from the Summary file to the HTML file
				intRowCount = Datatable.GetSheet(strResultSheet).GetRowCount
				For intRowCounter = 1 To intRowCount
					Datatable.GetSheet(strResultSheet).SetCurrentRow(intRowCounter)	
					strTC_ID = Datatable("TC_ID", strResultSheet)
					strDescription = Datatable("Description", strResultSheet)
					strExecutionTime = Datatable("Execution_Time_Minutes", strResultSheet)
					strStatus = Datatable("Status", strResultSheet)
					strLnkFileName = strTC_ID
					
					objMyFile.Writeline("<tr bgcolor = " & strHeadColor & ">")	

						objMyFile.Writeline("<td width=" & "400>")							
				        	objMyFile.Writeline("<p align=center><a href='" & strLnkFileName & ".html" & "'" & "target=" & "about_blank" & "><font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">"  & SheetVal)
						objMyFile.Writeline("</td>")
					   
						objMyFile.Writeline("<td width=" & "400>")							
				        	objMyFile.Writeline("<p align=center><a href='" & strLnkFileName & ".html" & "'" & "target=" & "about_blank" & "><font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">"  & strTC_ID)
						objMyFile.Writeline("</td>")
						
						objMyFile.Writeline("<td width=" & "400>")
					        objMyFile.Writeline("<p align=" & "center><font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">"  &  strDescription)
						objMyFile.Writeline("</td>")			
						
						objMyFile.Writeline("<td width=" & "400>")
					        objMyFile.Writeline("<p align=" & "center><font color=#000000 size=2 face= "& chr(34)&"Copperplate Gothic"&chr(34) & ">"  &  strExecutionTime)
						objMyFile.Writeline("</td>")		
						
						objMyFile.Writeline("<td width=" & "400>")
							If UCase(strStatus) = "PASS" Then
								objMyFile.Writeline("<p align=" & "center" & ">" & "<font face=" & "Verdana " & "size=" & "2" & " color=" & "#008000" & ">" & strStatus & "</font>")
								intPassCounter = intPassCounter + 1
							ElseIf UCase(strStatus) = "FAIL" Then
								objMyFile.Writeline("<p align=" & "center" & ">" & "<font face=" & "Verdana " & "size=" & "2" & " color=" & "#FF0000" & ">" & strStatus & "</font>")
								intFailCounter = intFailCounter + 1
							Else
								objMyFile.Writeline("<p align=" & "center" & ">" & "<font face=" & "Verdana " & "size=" & "2" & " color=" & "#8A4117" & ">" & strStatus & "</font>")
								intNoRunCounter=intNoRunCounter + 1
							End If
						objMyFile.Writeline("</td>")			
					objMyFile.Writeline("</tr>")	
				Next
				objMyFile.Writeline("</table>")
				
'				objMyFile.Writeline("<table border=1 bordercolor=" & "#000000 id=table1 width=900 height=31 cellspacing=0 bordercolorlight=" & "#FFFFFF>")	
'					objMyFile.Writeline("<tr bgcolor =" & strSettColor &">")
'						objMyFile.Writeline("<td colspan =1>")
'							objMyFile.Writeline("<p align=justify><b><font color=" & strHeadColor & " size=2 face= Verdana>"& "&nbsp;"& "Passed :&nbsp;&nbsp;" &  intPassCounter & "&nbsp;")
'						objMyFile.Writeline("</td>")
'						
'						objMyFile.Writeline("<td colspan =1>")	
'							objMyFile.Writeline("<p align=justify><b><font color=" & strHeadColor & " size=2 face= Verdana>"& "&nbsp;"& "Failed :&nbsp;&nbsp;" &  intFailCounter & "&nbsp;")
'						objMyFile.Writeline("</td>")
'						
'						objMyFile.Writeline("<td colspan =1>")	
'							objMyFile.Writeline("<p align=justify><b><font color=" & strHeadColor & " size=2 face= Verdana>"& "&nbsp;"& "InComplete :&nbsp;&nbsp;" &  intNoRunCounter & "&nbsp;")
'						objMyFile.Writeline("</td>")	
'					objMyFile.Writeline("</tr>")
'				objMyFile.Writeline("</table>")
			objMyFile.Writeline("</blockquote>")				
		objMyFile.Writeline("</body>")
	objMyFile.Writeline("</html>")
	objMyFile.Close
	
	Set objMyFile = Nothing
	Set objFso = Nothing 

	Dim strCurrentTestCase
	strCurrentTestCase = Environment.Value("CurrentTestCase")
	strPath = Environment.Value("ResultPath") & "\" &_
				Environment.Value("TimeStamp") & "\HTML Results\" &_
										strCurrentTestCase & ".html"	

	dim filesys
	Set filesys = CreateObject("Scripting.FileSystemObject")
	If filesys.FileExists(strPath) Then
		filesys.DeleteFile strPath
	End If 

	RACK_ExportReportedEventsToHtml strProjectName, strCurrentTestCase, strReportsTheme

End Sub
'#####################################################################################################################




'######################################################################################################################

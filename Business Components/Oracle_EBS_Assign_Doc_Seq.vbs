
'#####################################################################################################################
'Function Description   : Function to Navigate to assign Document Sequence
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Naviagte_Doc_Seq()
	'set WshShell = CreateObject("WScript.Shell")
	'set Parent_desc = Browser("Oracle Applications Home Page").Page("Suppliers")
	Select_Link(RACK_GetData("Assign_Doc_Seq_Data", "Responsibility_Link"))
	wait(10)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_6").Link("Assign").Click
	'wait(3)
	wait(60)
	OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Document").OracleTable("Table").EnterField 1,"Application",RACK_GetData("Assign_Doc_Seq_Data", "Application")
	wait(3)
	OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Document").OracleTable("Table").EnterField 1,"Category",RACK_GetData("Assign_Doc_Seq_Data", "Category")
	wait(3)
	OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Document").OracleTable("Table").EnterField 1,"Ledger",RACK_GetData("Assign_Doc_Seq_Data", "Ledger")
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(3)
	'OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Document").OracleTable("Table").EnterField 1,"Method",RACK_GetData("Assign_Doc_Seq_Data", "Method")
	OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Document").OracleTable("Table").EnterField 1,"Method","Null"
    'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(3)
	OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Assignment").OracleTable("Table").EnterField 1,"Start Date",RACK_GetData("Assign_Doc_Seq_Data", "Start_Date")
	'OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Assignment").OracleTable("Table").EnterField 1,"Start Date","01-JAN-1960"
	If not RACK_GetData("Assign_Doc_Seq_Data", "End_Date") = "" Then
	OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Assignment").OracleTable("Table").EnterField 1,"End Date",RACK_GetData("Assign_Doc_Seq_Data", "End_Date")
	End If
	OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Assignment").OracleTable("Table").EnterField 1,"Sequence",RACK_GetData("Assign_Doc_Seq_Data", "Sequence")
	wait(3)
	OracleFormWindow("Sequence Assignments").PressToolbarButton "Clear Record"
	OracleFormWindow("Sequence Assignments").SelectMenu "File->Save"
	wait(2)
	Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
End Function

'#####################################################################################################################
'Function Description   : Function to Assign Document Sequence
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Assign_Doc_Seq()
		OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Document").OracleTable("Table").SetFocus 1,"Application"
		wait(2)
		OracleFormWindow("Sequence Assignments").SelectMenu "File->New"
		wait(2)
		OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Document").OracleTable("Table").EnterField 2,"Application",RACK_GetData("Assign_Doc_Seq_Data", "Application")
		wait(2)
		OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Document").OracleTable("Table").EnterField 2,"Category",RACK_GetData("Assign_Doc_Seq_Data", "Category")
		wait(2)
		OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Document").OracleTable("Table").EnterField 2,"Ledger",RACK_GetData("Assign_Doc_Seq_Data", "Ledger")
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(2)
		'OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Document").OracleTable("Table").EnterField 2,"Method",RACK_GetData("Assign_Doc_Seq_Data", "Method")
		OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Document").OracleTable("Table").EnterField 1,"Method","Null"
        'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(2)
		OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Assignment").OracleTable("Table").EnterField 2,"Start Date",RACK_GetData("Assign_Doc_Seq_Data", "Start_Date")
		If not RACK_GetData("Assign_Doc_Seq_Data", "End_Date") = "" Then
	    OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Assignment").OracleTable("Table").EnterField 2,"End Date",RACK_GetData("Assign_Doc_Seq_Data", "End_Date")
		End If
		OracleFormWindow("Sequence Assignments").OracleTabbedRegion("Assignment").OracleTable("Table").EnterField 2,"Sequence",RACK_GetData("Assign_Doc_Seq_Data", "Sequence")
		wait(2)
		'OracleFormWindow("Sequence Assignments").PressToolbarButton "Clear Record"
		OracleFormWindow("Sequence Assignments").SelectMenu "File->Save"
		wait(2)
		Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
End Function
'########################################################################################################################



'########################################################################################################################

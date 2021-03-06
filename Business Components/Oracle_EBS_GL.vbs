'#####################################################################################################################
'Function Description   : Function to Validate all subledger transactions from AR, AP
'Input Parameters 	: None
'Return Value    	: None
'Created By :  Ram
'##################################################################################################################### 
Function Validate_All_Subledger_Transactions_Froms()
If Browser("Oracle Applications Home Page").Page("Oracle Applications Home").Link("RS US GL User").Exist(60) Then
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home").Link("RS US GL User").Click
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home").Link("Enter").Click
	wait(60)
		If not RACK_GetData("GL_Data", "Journel_Source") = "" Then
		OracleFormWindow("Find Journals").OracleTextField("Source").Enter RACK_GetData("GL_Data","Journel_Source")
		End If
		If not RACK_GetData("GL_Data", "Journel_Category") = "" Then
		OracleFormWindow("Find Journals").OracleTextField("Category").Enter RACK_GetData("GL_Data","Journel_Category")
		End If
		OracleFormWindow("Find Journals").OracleTextField("Period").Enter RACK_GetData("GL_Data","Period")
		RACK_ReportEvent "Validation Screenshot", "Find Journals values Successfully Enter ","Screenshot"
		wait(2)
		OracleFormWindow("Find Journals").OracleButton("Find").Click
		wait(10)
		RACK_ReportEvent "Validation Screenshot", "Find Journals Page Open ","Screenshot"
		OracleFormWindow("Enter Journals").OracleButton("Review Journal").Click
		wait(2)
		RACK_ReportEvent "Validation Screenshot", "Journal Account Page sucessfully Open ","Pass"
		End If
End Function
'#####################################################################################################################
'Function Description   : Function for Post  Journals
'Input Parameters 	: None
'Return Value    	: None
'Created By :  Avvaru Nagarjuna
'##################################################################################################################### 
Function Post_Unposted_AR_journals()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("GL_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("GL_Data", "Functionality_Link"))

		If OracleFormWindow("Find Journal Batches").Exist(120) Then
			If not RACK_GetData("GL_Data", "Journel_Batch_Name") = "" Then
				Query = RACK_GetData("GL_Data", "Journel_Batch_Name") & "%"
				OracleFormWindow("Find Journal Batches").OracleTextField("Batch").Enter Query
				'OracleFormWindow("Find Journal Batches").OracleTextField("Batch").Enter Query
            	OracleFormWindow("Find Journal Batches").OracleTextField("Period").Enter RACK_GetData("GL_Data", "Period")
			End If
			wait(2)
			OracleFormWindow("Find Journal Batches").OracleButton("Find").Click
			If  OracleFormWindow("Post Journals").Exist(10) Then
                OracleFormWindow("Post Journals").OracleTable("Table").EnterField 1,"Selected for Posting", TRUE
				wait(2)
				OracleFormWindow("Post Journals").OracleButton("Post").Click
                RACK_ReportEvent "Validation Screenshot", "Journal successfully posted   ","Screenshot"
				Notification_Message = Get_Oracle_Notification_Form_Message()
				Arr=split(Notification_Message, " ")
				Posting_ID_Arr = Split(Arr(6),".")
				Posting_ID = Posting_ID_Arr(0)
				Handle_Oracle_Notification_Forms("OK")
			End If
			OracleFormWindow("Find Journal Batches").CloseWindow
			wait(2)
			OracleFormWindow("Post Journals").CloseWindow
			wait(2)
			OracleFormWindow("Navigator").SelectMenu "View->Requests"
			OracleFormWindow("Find Requests").OracleRadioGroup("My Completed Requests").Select "Specific Requests"
			OracleFormWindow("Find Requests").OracleTextField("Request ID").Enter Posting_ID
			wait(2)
			OracleFormWindow("Find Requests").OracleButton("Find").Click
			wait(2)
			OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			wait(2)
            OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			wait(2)
            OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			wait(2)
            OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			wait(2)
            OracleFormWindow("Requests").OracleButton("Refresh Data").Click
 Phase = OracleFormWindow("Requests").OracleTable("Table").GetFieldValue(1,"Phase")
 If Phase = "Completed" Then
 RACK_ReportEvent "Verification Post Journel -Phase is Completed","Phase = Completed","Pass"
 Else
 RACK_ReportEvent "Verification Post Journe -Phase is Completed","Phase is - "&Phase,"Fail"
			end if
		end if
End Function
'#####################################################################################################################
'Function Description   : Function for manual journal entry
'Input Parameters 	: None
'Return Value    	: None
'Created By :  Ram
'##################################################################################################################### 
Function Manual_Journal_Entry()
If Browser("Oracle Applications Home Page").Page("Oracle Applications Home").Link("RS US GL User").Exist(60) Then
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home").Link("RS US GL User").Click
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home").Link("Enter").Click
	wait(5)
	If OracleFormWindow("Find Journals").Exist(120)  Then
		OracleFormWindow("Find Journals").OracleButton("New Journal").Click
		wait(2)
		If OracleFormWindow("Journals").Exist(60)  Then
			OracleFormWindow("Journals").OracleTextField("Journal").Enter RACK_GetData("GL_Data","Journel_Name")
			wait(2)
			If not  RACK_GetData("GL_Data", "Description") ="" Then
			OracleFormWindow("Journals").OracleTextField("Description").Enter RACK_GetData("GL_Data","Description")
			'End If
			wait(2)
			'OracleFormWindow("Journals").OracleTextField("Period").SetFocus
			'OracleFormWindow("Journals").OracleTextField("Reverse|Period").Enter RACK_GetData("GL_Data","Period")
			OracleFormWindow("Journals").OracleTextField("Period").Enter RACK_GetData("GL_Data","Period")
			wait(2)
			OracleFormWindow("Journals").OracleTextField("Category").Enter RACK_GetData("GL_Data","Journel_Category")
			wait(2)
			OracleFormWindow("Journals").OracleTabbedRegion("Lines").OracleTable("Table").SetFocus 1,"Line"
            OracleFormWindow("Journals").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Line",RACK_GetData("GL_Data","Journel_Line1")
			wait(2)
			OracleFormWindow("Journals").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Account",RACK_GetData("GL_Data","JournelLine_Account")
				wait(2)
			OracleFormWindow("Journals").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Debit (USD)",RACK_GetData("GL_Data","Debit")
				wait(2)
			OracleFormWindow("Journals").OracleTabbedRegion("Lines").OracleTable("Table_2").SetFocus 2,"Line"
				wait(2)
			OracleFormWindow("Journals").OracleTabbedRegion("Lines").OracleTable("Table_2").EnterField 2,"Line",RACK_GetData("GL_Data","Journel_Line2")
				wait(2)
			OracleFormWindow("Journals").OracleTabbedRegion("Lines").OracleTable("Table_2").EnterField 2,"Account",RACK_GetData("GL_Data","JournelLine_Account")
				wait(2)
			OracleFormWindow("Journals").OracleTabbedRegion("Lines").OracleTable("Table_2").EnterField 2,"Credit (USD)",RACK_GetData("GL_Data","Credit")
			wait(2)
			OracleFormWindow("Journals").SelectMenu "File->Save"
			wait(2)
			RACK_ReportEvent "Verification -Transaction Complete","Verification = Transaction Complete 3 records applied and saved","Pass"
		End If
	End If
End If
End If
End Function
'#####################################################################################################################
''Function Description   : Function to Query specific account to check activity
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 
Function Query_Specific_Account()
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("GL_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("GL_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Account Inquiry").OracleTextField("Accounting Periods|From").Enter RACK_GetData("GL_Data","Period")
	wait(2)
	OracleFormWindow("Account Inquiry").OracleTextField("Accounting Periods|To").Enter RACK_GetData("GL_Data","Period")
	wait(2)
	OracleFormWindow("Account Inquiry").OracleTable("Table").SetFocus 1,"Account"
	wait(2)
	OracleFlexWindow("Find Accounts").OracleTextField("Company").Enter "100"
	wait(2)
	OracleFlexWindow("Find Accounts").OracleTextField("Location").Enter "000"
	wait(2)
	OracleFlexWindow("Find Accounts").OracleTextField("Account").Enter "159990"
	wait(2)
	OracleFlexWindow("Find Accounts").OracleTextField("Team").Enter "000"
	wait(2)
    OracleFlexWindow("Find Accounts").OracleTextField("Business Unit").Enter "0000"
	wait(2)
	OracleFlexWindow("Find Accounts").OracleTextField("Department").Enter "0000"
	wait(2)
	OracleFlexWindow("Find Accounts").OracleTextField("Product").Enter "0000"
	wait(2)
	OracleFlexWindow("Find Accounts").OracleTextField("Future").Enter "0000"
    wait(2)
	RACK_ReportEvent "Validation Screenshot", "Find Accounts Parameters successfully Enter   ","Screenshot"
	OracleFlexWindow("Find Accounts").Approve
	wait(2)
	OracleFormWindow("Account Inquiry").OracleButton("Show Balances").Click
	wait(2)
	RACK_ReportEvent "Validation Screenshot", "Journal Balances should be populated Sucessfully ","Screenshot"
	wait(2)
	OracleFormWindow("Detail Balances").CloseWindow
	wait(2)
	OracleFormWindow("Account Inquiry").OracleButton("Show Balances").Click
	wait(2)
	OracleFormWindow("Detail Balances").OracleButton("Journal Details").Click
	RACK_ReportEvent "Validation Screenshot", "Journal Detail Balances should be populated Sucessfully ","Screenshot"
End Function
'#####################################################################################################################
''Function Description   : Function to Run Trial Balance Summary-2
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 
Function Run_TrialBalance_Summary_II()
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("GL_Data", "Responsibility_Link"))
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_39").Link("Run").Click
	wait(60)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	Wait(5)
	OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("GL_Data","Request_Name")
	wait(2)
OracleFlexWindow("Parameters").OracleTextField("Balance Type").Enter RACK_GetData("GL_Data","Balance_Type")
wait(2)
OracleFlexWindow("Parameters").OracleTextField("Pagebreak Segment").Enter RACK_GetData("GL_Data","Segment_Value")
wait(2)
OracleFlexWindow("Parameters").OracleTextField("Pagebreak Segment Low").SetFocus
OracleFlexWindow("Accounting Flexfield").OracleTextField("Company_2").Enter RACK_GetData("GL_Data","Company_Number")
OracleFlexWindow("Accounting Flexfield").Approve
wait(5)
OracleFlexWindow("Parameters").OracleTextField("Pabebreak Segment High").SetFocus
OracleFlexWindow("Parameters").OracleTextField("Secondary Segment").Enter RACK_GetData("GL_Data","Secondary_Segment")
wait(5)
OracleFlexWindow("Parameters").OracleTextField("Period").Enter RACK_GetData("GL_Data","Period")
wait(2)
OracleFlexWindow("Parameters").OracleTextField("Budget Start Period").Enter RACK_GetData("GL_Data","Budget_Type")
wait(2)
OracleFlexWindow("Parameters").OracleTextField("Amount Type").Enter RACK_GetData("GL_Data","Amount_Type")
wait(2)
'RACK_ReportEvent "Validation Screenshot", "Trial_Balance_Summary Parameters successfully Enter   ","Screenshot"
OracleFlexWindow("Parameters").Approve
wait(2)
	OracleFormWindow("Submit Request").OracleButton("Submit").Click
    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	Notification_Message = Get_Oracle_Notification_Form_Message()
	Arr=split(Notification_Message, " ")
	Content = Arr(6)
	Payment_Request_ID_Arr = Split(Content,")")
	Payment_Request_ID = Payment_Request_ID_Arr(0)
	Update_Notepad "Payment_Request_ID", Payment_Request_ID
	Handle_Oracle_Notification_Forms("No")
	Verify_Request_Status(Payment_Request_ID)
End Function

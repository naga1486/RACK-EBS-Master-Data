
'#####################################################################################################################
'CPU Patch Testing
'Function Description   : Function to Create a new employee
'Input Parameters 	: None
'Return Value    	: None
'Created By Avvaru Nagarjuna
'##################################################################################################################### 
Function Create_New_Employee()
	Select_Link(RACK_GetData("HR_Data", "Responsibility_Link"))
	wait(10)
	Select_Link(RACK_GetData("HR_Data", "Functionality_Link"))
		If OracleFormWindow("Find Person").Exist(90) Then
			OracleFormWindow("Find Person").OracleButton("New").Click
			If OracleFormWindow("People").Exist(10) Then
				OracleFormWindow("People").OracleTextField("Name|Last").Enter RACK_GetData("HR_Data", "Employee_Last_Name")
				wait(2)
				OracleFormWindow("People").OracleTextField("Name|First").Enter RACK_GetData("HR_Data", "Employee_First_Name")
				wait(2)
				OracleFormWindow("People").OracleList("Gender").Select RACK_GetData("HR_Data", "Employee_Gender")
				wait(2)
				OracleFormWindow("People").OracleList("Action").Select RACK_GetData("HR_Data", "Employee_Action")
				wait(2)			
				ssn = CStr(Int((999-100+1)*Rnd+100)) &"-"&CStr(Int((99-10+1)*Rnd+10))&"-"&CStr(Int((9999-1000+1)*Rnd+1000))
				OracleFormWindow("People").OracleTextField("Identification|Social").Enter ssn
				OracleFormWindow("People").OracleTabbedRegion("Personal").OracleTextField("Birth Date").Enter RACK_GetData("HR_Data", "Employee_DOB")
				OracleFormWindow("People").OracleTextField("[").SetFocus
				wait(2)
				OracleFlexWindow("Additional Personal Details").OracleTextField("Workday Employee ID").Enter RACK_GetData("HR_Data", "Workday_Emp_Id")
				OracleFlexWindow("Additional Personal Details").Approve
				If OracleFormWindow("Choose an option").Exist(10) Then
					OracleFormWindow("Choose an option").OracleButton("Correction").Click
				End If
				wait(2)
				OracleFormWindow("People").SelectMenu "File->Save"
				wait(5)
				status = OracleStatusLine("OracleStatusLine").GetROProperty("message")
				If  InStr(1,status,"saved") > 0 Then
				RACK_ReportEvent "Create Employee", "The Employee '" & RACK_GetData("HR_Data", "Employee_Last_Name") & "'  has been sucessfully Created" ,"Screenshot"
                'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				Else
				RACK_ReportEvent " Create Employee", "Employee not created successfully" ,"Fail"
				End If
			End If
		End If
End Function
'#####################################################################################################################
'Function Description   : Function to Assignments 
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Assignments()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Suppliers")
	Select_Link(RACK_GetData("HR_Data", "Responsibility_Link"))
	wait(10)
	Select_Link(RACK_GetData("HR_Data", "Functionality_Link"))
	If OracleFormWindow("Find Person").Exist(120) Then
		OracleFormWindow("Find Person").OracleTextField("Full Name").Enter RACK_GetData("HR_Data", "Employee_Full_Name")
		OracleFormWindow("Find Person").OracleButton("Find").Click
		wait(2)
		If OracleFormWindow("People").Exist(20) Then
			OracleFormWindow("People").OracleButton("Assignment").Click
			If  OracleFormWindow("Assignment").Exist(15) Then
				OracleFormWindow("Assignment").OracleTabbedRegion("Supervisor").OracleTextField("Name").Enter RACK_GetData("HR_Data", "Employee_Supervisor_Name")
				OracleFormWindow("Assignment").SelectMenu "File->Save"	
				wait(2)
				If  OracleFormWindow("Choose an option").Exist(5) Then
					OracleFormWindow("Choose an option").OracleButton("Correction").Click
					End If
				OracleFormWindow("Assignment").OracleTextField("Job").Enter RACK_GetData("HR_Data", "Employee_Job")
				RACK_ReportEvent "Validation Screenshot", "Job  Assignment successfully Enter   ","Screenshot"
				wait(2)
				If  OracleFormWindow("Choose an option").Exist(5) Then
					OracleFormWindow("Choose an option").OracleButton("Correction").Click
				End If
                wait(2)
                OracleFormWindow("Assignment").OracleTabbedRegion("Purchase Order Information").Click'
				wait(2)
				OracleFormWindow("Assignment").OracleTabbedRegion("Purchase Order Information").OracleTextField("Ledger").SetFocus
				wait(2)
				OracleFormWindow("Assignment").OracleTabbedRegion("Purchase Order Information").OracleTextField("Ledger").Enter RACK_GetData("HR_Data", "Employee_Ledger")
				OracleFlexWindow("Accounting Flexfield").OracleTextField("Company").Enter RACK_GetData("HR_Data", "Employee_POCompany")
				wait(2)
				OracleFlexWindow("Accounting Flexfield").OracleTextField("Account").Enter RACK_GetData("HR_Data", "Employee_POAccount")
				wait(2)
				RACK_ReportEvent "Validation Screenshot", "Po Accounts Parameters successfully Enter   ","Screenshot"
				OracleFlexWindow("Accounting Flexfield").Approve
				wait(2)
				OracleFormWindow("Assignment").SelectMenu "File->Save"
				wait(2)
				Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
				status = OracleStatusLine("OracleStatusLine").GetROProperty("message")
				If  InStr(1,status,"saved") > 0 Then
					RACK_ReportEvent "Update Employee", "The Employee '" & RACK_GetData("HR_Data", "Employee_Full_Name") & "'  has been sucessfully Updated" ,"Pass"
				Else
					RACK_ReportEvent "Update Employee", "Employee not updated successfully" ,"Fail"
				End If
			End If
		End If
	End If
End Function
'#####################################################################################################################
'Function Description   : Function to Terminate an Employee 
'Input Parameters 	: None
'Return Value    	: None
'Created By : Avvaru Nagarjuna
'##################################################################################################################### 
Function Terminate_Employee()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Suppliers")
	Select_Link(RACK_GetData("HR_Data", "Responsibility_Link"))
	wait(10)
	Select_Link(RACK_GetData("HR_Data", "Functionality_Link"))
	If OracleFormWindow("Find Person").Exist(120) Then
		OracleFormWindow("Find Person").OracleTextField("Full Name").Enter RACK_GetData("HR_Data", "Employee_Full_Name")
		OracleFormWindow("Find Person").OracleButton("Find").Click
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(2)
		OracleFormWindow("People").OracleButton("&Others...").Click
		wait(2)
		OracleListOfValues("Navigation Options").Select "End Employment"
        wait(5)
		OracleFormWindow("Terminate").OracleTextField("Leaving Reason").Enter RACK_GetData("HR_Data", "Leaving_Reason")
		wait(2)
		OracleFormWindow("Terminate").OracleTextField("Termination Dates|Notified").Enter RACK_GetData("HR_Data", "Date_Notified")
		wait(2)
        OracleFormWindow("Terminate").OracleTextField("Termination Dates|Projected").Enter RACK_GetData("HR_Data", "Date_Projected")
		wait(2)
        OracleFormWindow("Terminate").OracleTextField("Termination Dates|Actual").Enter RACK_GetData("HR_Data", "Date_Actual")
		wait(10)
		OracleFormWindow("Terminate").OracleTextField("Termination Accepted By|Date").SetFocus
		wait(2)
		OracleFormWindow("Terminate").OracleTextField("Termination Accepted By|Date").Enter RACK_GetData("HR_Data", "Accepted_By_Date")
		wait(2)
        OracleFormWindow("Terminate").OracleTextField("Termination Accepted By|Name").Enter RACK_GetData("HR_Data", "Accepted_By_Name")
		wait(2)
		OracleFormWindow("Terminate").OracleButton("Terminate").Click
		Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
		wait(10)
		OracleFormWindow("Terminate").CloseWindow
		wait(2)
		OracleFormWindow("People").SelectMenu "View->Find..."
		wait(5)
		OracleFormWindow("Find Person").OracleTextField("Full Name").Enter RACK_GetData("HR_Data", "Employee_Full_Name")
		wait(2)
		OracleFormWindow("Find Person").OracleButton("Find").Click
	    wait(5)
	   RACK_ReportEvent "Validation Screenshot", "Employee is sucessfully Terminated ","Pass"
	   End If
End Function
'#####################################################################################################################
'Function Description   : Function to Run Duplicate person report
'Input Parameters 	: None
'Return Value    	: None
'Created By :  Avvaru Nagarjuna
'##################################################################################################################### 
Function Run_Duplicate_Person_Report()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Suppliers")
	Select_Link(RACK_GetData("HR_Data", "Responsibility_Link"))
	wait(10)
	Select_Link(RACK_GetData("HR_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Find Person").SelectMenu "View->Requests"
	wait(2)
    OracleFormWindow("Find Requests").OracleButton("Submit a New Request...").Click
	wait(2)
    OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	If  OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Exist(30) Then
		OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("HR_Data","Request_Name")
		wait(2)
		OracleFlexWindow("Parameters").Approve
		wait(2)
		OracleFormWindow("Submit Request").OracleButton("Submit").Click
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	End If
		Notification_Message = Get_Oracle_Notification_Form_Message()
			Arr=split(Notification_Message, " ")
			Content = Arr(6)
			Misc_Payment_Request_ID_Arr = Split(Content,")")
			Misc_Payment_Request_ID = Misc_Payment_Request_ID_Arr(0)
			Update_Notepad "Misc_Payment_Request_ID", Misc_Payment_Request_ID
			Handle_Oracle_Notification_Forms("No")
			wait(2)
			OracleFormWindow("Find Requests").CloseWindow
			wait(2)
			Verify_Request_Status(Misc_Payment_Request_ID)
End Function
'#####################################################################################################################
'Function Description   : Function to Modify an existing Employee Record
'Input Parameters 	: None
'Return Value    	: None
'Created By :  Avvaru Nagarjuna
'##################################################################################################################### 
Function Modify_Employee_Record()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Suppliers")
	Select_Link(RACK_GetData("HR_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("HR_Data", "Functionality_Link"))
	If OracleFormWindow("Find Person").Exist(120) Then
		OracleFormWindow("Find Person").OracleTextField("Full Name").Enter RACK_GetData("HR_Data", "Employee_Full_Name")
		OracleFormWindow("Find Person").OracleButton("Find").Click
			wait(2)
			If OracleFormWindow("People").Exist(20)  Then
				OracleFormWindow("People").OracleButton("Assignment").Click
				wait(2)
				If OracleFormWindow("Assignment").Exist(20)  Then
					OracleFormWindow("Assignment").OracleTextField("Assignment Category").Enter RACK_GetData("HR_Data", "Assignment_Category")
					wait(2)
					If OracleFormWindow("Choose an option").Exist(20)  Then
						OracleFormWindow("Choose an option").OracleButton("Update").Click
						wait(2)
						If OracleNotification("Error").Exist(10)  Then
							OracleNotification("Error").OracleButton("OK").Click
							wait(2)
							OracleFormWindow("Assignment").OracleTextField("Job").Enter RACK_GetData("HR_Data", "Employee_Job")
						If OracleFormWindow("Assignment").Exist(20)  Then
							OracleFormWindow("Assignment").OracleTextField("Location").Enter RACK_GetData("HR_Data", "Location")					
							OracleFormWindow("Assignment").SelectMenu "File->Save"
								wait(2)
					            status = OracleStatusLine("OracleStatusLine").GetROProperty("message")
					If  InStr(1,status,"saved") > 0 Then
						RACK_ReportEvent "Update Assignment", "The Employee '" & RACK_GetData("HR_Data", "Employee_Full_Name") & "'  has been sucessfully Updated" ,"Pass"
					Else
						RACK_ReportEvent "Update Assignment", "Employee not updated successfully" ,"Fail"
					  End If
						End If
					End If
				End If
			End If
		End If
	End If
End Function
'#####################################################################################################################################



'#####################################################################################################################################

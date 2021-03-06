'#####################################################################################################################
'Function Description   : Function to Workday AP Expenses
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Workday_AP_Expenses()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Workday_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("Workday_Data", "Functionality_Link"))
    If OracleFormWindow("Submit a New Request").Exist(90) Then
		OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request Set").Exist(10) Then
			OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("Workday_Data", "Request_Name")
			wait(2)
			If not RACK_GetData("Workday_Data", "File_Name") = "" Then
				OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
				OracleFlexWindow("Parameters").OracleTextField("File Name").Enter RACK_GetData("Workday_Data", "File_Name")
				 RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				OracleFlexWindow("Parameters").Approve
			End If
			If not RACK_GetData("Workday_Data", "Operating_Unit") = ""  OR not RACK_GetData("Workday_Data", "Supplier_Name") = ""Then
				       OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 3,"Parameters"
				wait(2)
				OracleFlexWindow("Parameters").OracleTextField("Operating Unit").Enter RACK_GetData("Workday_Data", "Operating_Unit")
			If not RACK_GetData("Workday_Data", "Option") = "" Then
					OracleFlexWindow("Parameters").OracleTextField("Option").Enter RACK_GetData("Workday_Data", "Option")
				End If		
				OracleFlexWindow("Parameters").OracleTextField("From Invoice Date").Enter RACK_GetData("Workday_Data", "From_Inv_Date")
				OracleFlexWindow("Parameters").OracleTextField("To Invoice Date").Enter RACK_GetData("Workday_Data", "To_Inv_Date")
				OracleFlexWindow("Parameters").OracleTextField("Supplier Name").Enter RACK_GetData("Workday_Data", "Supplier_Name")
				  RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				OracleFlexWindow("Parameters").Approve
			End If
			wait(2)
			OracleFormWindow("Submit Request Set").OracleButton("Submit").Click
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			Handle_Oracle_Notification_Forms("OK")
			wait(2)
			Notification_Message = Get_Oracle_Notification_Form_Message()
			Arr=split(Notification_Message, " ")
			Content = Arr(6)
			Print_Request_ID_Arr = Split(Content,")")
			Print_Request_ID = Print_Request_ID_Arr(0)
			Update_Notepad "Print_Request_ID", Print_Request_ID
			Handle_Oracle_Notification_Forms("No")
			Verify_Request_Status(Print_Request_ID)
		else
			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
		end if

	else
			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	end if
End Function					

'#####################################################################################################################
'Function Description   : Function to Workday GL Expenses
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Workday_GL_Expenses()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Workday_Data", "Responsibility_Link"))
	wait(3)
	'Select_Link(RACK_GetData("Workday_Data", "Functionality_Link"))
	Parent_desc.Link("Run").Click
	If OracleFormWindow("Submit a New Request").Exist(90) Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(10) Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Workday_Data", "Request_Name")
			wait(2)
			If OracleFlexWindow("Parameters").Exist(4)  Then
				OracleFlexWindow("Parameters").OracleButton("OK").Click
				wait(5)
			end if
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			Handle_Oracle_Notification_Forms("OK")
			wait(2)
			Notification_Message = Get_Oracle_Notification_Form_Message()
			Arr=split(Notification_Message, " ")
			Content = Arr(6)
			Print_Request_ID_Arr = Split(Content,")")
			Print_Request_ID = Print_Request_ID_Arr(0)
			Update_Notepad "Print_Request_ID", Print_Request_ID
			Handle_Oracle_Notification_Forms("No")
			Verify_Request_Status(Print_Request_ID)
		else
			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
		end if
	else
			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	end if
End Function
'######################################################################################################################



'######################################################################################################################

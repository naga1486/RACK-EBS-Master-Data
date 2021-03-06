'#####################################################################################################################
'CPU Patch Testing
'Function Description   : Function to Define User 
'Input Parameters 	: None
'Return Value    	: None
'Created By : Avvaru Nagarjuna
'##################################################################################################################### 
Function Define_User()
   	Select_Link(RACK_GetData("Sys_Admin_Data", "Responsibility_Link"))
	'Browser("Oracle Applications Home").Page("Oracle Applications Home_33").Link("System Administrator").Click
	wait(10)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_31").Link("Define").Click
	wait(60)
	OracleFormWindow("Users").OracleTextField("User Name").Enter RACK_GetData("Sys_Admin_Data", "User_Name")
	wait(2)
    OracleFormWindow("Users").OracleTextField("Password").Enter RACK_GetData("Sys_Admin_Data", "Password")
    OracleFormWindow("Users").SelectMenu "File->Save"
    wait(2)
    OracleFormWindow("Users").OracleTextField("Password").Enter RACK_GetData("Sys_Admin_Data", "Password")
    OracleFormWindow("Users").SelectMenu "File->Save"
	Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
	wait(5)
	OracleFormWindow("Users").OracleTabbedRegion("Direct Responsibilities").OracleTable("Table").Setfocus 1,"Responsibility"
	wait(5)
	OracleFormWindow("Users").OracleTabbedRegion("Direct Responsibilities").OracleTable("Table").EnterField 1,"Responsibility",RACK_GetData("Sys_Admin_Data","Responsibility")
	wait(2)
	OracleFormWindow("Users").SelectMenu "File->Save"
	Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
End Function
'#####################################################################################################################
'Function Description   : Function to Assign Responsibility to User
'Input Parameters 	: None
'Return Value    	: None
'Created By : Avvaru Nagarjuna
'##################################################################################################################### 
Function Assign_Responsibility_To_User()
   	Select_Link(RACK_GetData("Sys_Admin_Data", "Responsibility_Link"))
	wait(3)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_31").Link("Define").Click
	wait(60)
    OracleFormWindow("Users").SelectMenu "View->Query By Example->Enter"
	wait(3)
    OracleFormWindow("Users").OracleTextField("User Name").Enter RACK_GetData("Sys_Admin_Data", "User_Name")
	wait(3)
	RACK_ReportEvent "Validation Screenshot", "User Name Successfully Enter   ","Screenshot"
    OracleFormWindow("Users").SelectMenu "View->Query By Example->Run"
	wait(3)
    OracleFormWindow("Users").OracleTabbedRegion("Direct Responsibilities").OracleTable("Table").Setfocus 1,"Responsibility"
	wait(2)
	OracleFormWindow("Users").SelectMenu "File->New"
	wait(2)
	OracleFormWindow("Users").OracleTabbedRegion("Direct Responsibilities").OracleTable("Table").EnterField 2,"Responsibility",RACK_GetData("Sys_Admin_Data","Responsibility")
	wait(2)
	OracleFormWindow("Users").SelectMenu "File->Save"
	Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
End Function
'#####################################################################################################################
'Function Description   : Function to End date Responsibility
'Input Parameters 	: None
'Return Value    	: None
'Created By : Avvaru Nagarjuna
'##################################################################################################################### 
Function End_Date_Responsibility()
   	Select_Link(RACK_GetData("Sys_Admin_Data", "Responsibility_Link"))
	wait(10)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_31").Link("Define").Click
	wait(60)
    OracleFormWindow("Users").SelectMenu "View->Query By Example->Enter"
	wait(3)
    OracleFormWindow("Users").OracleTextField("User Name").Enter RACK_GetData("Sys_Admin_Data", "User_Name")
	wait(3)
    OracleFormWindow("Users").SelectMenu "View->Query By Example->Run"
	wait(3)
	OracleFormWindow("Users").OracleTabbedRegion("Direct Responsibilities").OracleTable("Table").Setfocus 2,"To"
	wait(3)
	OracleFormWindow("Users").OracleTabbedRegion("Direct Responsibilities").OracleTable("Table").EnterField 2,"To",RACK_GetData("Sys_Admin_Data","End_Date")
	'OracleFormWindow("Users").OracleTabbedRegion("Direct Responsibilities").OracleTable("Table").Enter RACK_GetData("Sys_Admin_Data", "End_Date")
    OracleFormWindow("Users").SelectMenu "File->Save"
 Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
 wait(5)
End Function
'#####################################################################################################################
'Function Description   : Function to Un-End Date Responsibility 
'Input Parameters 	: None
'Return Value    	: None
'Created By : Avvaru Nagarjuna
'##################################################################################################################### 
Function Un_End_Date_Responsibility()
   	Select_Link(RACK_GetData("Sys_Admin_Data", "Responsibility_Link"))
	wait(3)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_31").Link("Define").Click
	wait(60)
    OracleFormWindow("Users").SelectMenu "View->Query By Example->Enter"
	wait(3)
    OracleFormWindow("Users").OracleTextField("User Name").Enter RACK_GetData("Sys_Admin_Data", "User_Name")
	wait(3)
    OracleFormWindow("Users").SelectMenu "View->Query By Example->Run"
	wait(5)
	OracleFormWindow("Users").OracleTabbedRegion("Direct Responsibilities").OracleTable("Table").Setfocus 2,"To"
	wait(3)
	OracleFormWindow("Users").OracleTabbedRegion("Direct Responsibilities").OracleTable("Table").EnterField 2,"To",""
	wait(3)
    OracleFormWindow("Users").SelectMenu "File->Save"
	Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
 End Function
'#####################################################################################################################
'Function Description   : Function to Run Function Security Reports
'Input Parameters 	: None
'Return Value    	: None
'Created By : Avvaru Nagarjuna
'##################################################################################################################### 
Function Function_Security_Reports()
   	Select_Link(RACK_GetData("Sys_Admin_Data", "Responsibility_Link"))
	wait(10)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_32").Link("Requests").Click
	wait(60)
	OracleFormWindow("Find Requests").OracleButton("Submit a New Request...").Click
	wait(2)
	OracleFormWindow("Submit a New Request").OracleRadioGroup("Request Set").Select "Request Set"
	wait(5)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("Sys_Admin_Data","Request_Name")
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Application").Enter RACK_GetData("Sys_Admin_Data", "Application")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Responsibility").Enter RACK_GetData("Sys_Admin_Data", "Responsibility")
	wait(2)
    OracleFlexWindow("Parameters").Approve
	wait(2)
    OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 2,"Parameters"
	wait(2)
    OracleFlexWindow("Parameters").Approve
	wait(2)
    OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 3,"Parameters"
	wait(2)
    OracleFlexWindow("Parameters").Approve
	wait(2)
	OracleFormWindow("Submit Request Set").OracleButton("Submit").Click
	wait(2)
	RACK_ReportEvent "Validation Screenshot", "Function Security Reports Parameters successfully Enter   ","Screenshot"
	wait(2)
	Notification_Message = Get_Oracle_Notification_Form_Message()
	Arr=split(Notification_Message, " ")
	Content = Arr(6)
	Payment_Request_ID_Arr = Split(Content,")")
	Payment_Request_ID = Payment_Request_ID_Arr(0)
	Update_Notepad "Payment_Request_ID", Payment_Request_ID
	Handle_Oracle_Notification_Forms("No")
	wait(2)
	OracleFormWindow("Find Requests").CloseWindow
	Verify_Request_Status(Payment_Request_ID)
End Function
'#####################################################################################################################
'Function Description   : Function to Run Synchronize WF LOCAL tables
'Input Parameters 	: None
'Return Value    	: None
'Created By : Avvaru Nagarjuna
'##################################################################################################################### 
Function Synchronize_WF_Local_Tables()
   	Select_Link(RACK_GetData("Sys_Admin_Data", "Responsibility_Link"))
	wait(10)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_32").Link("Requests").Click
	wait(60)
	OracleFormWindow("Find Requests").OracleButton("Submit a New Request...").Click
	wait(2)
	OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
	wait(2)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("Sys_Admin_Data","Request_Name")
	wait(2)
    OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
    OracleFlexWindow("Parameters").Approve
	wait(2)
    OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 2,"Parameters"
    OracleFlexWindow("Parameters").Approve
	wait(2)
    OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 3,"Parameters"
    OracleFlexWindow("Parameters").Approve
	wait(2)
    OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 4,"Parameters"
    OracleFlexWindow("Parameters").Approve
	wait(2)
    OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 5,"Parameters"
	wait(2)
    OracleFlexWindow("Parameters").Approve
	wait(2)
	OracleFormWindow("Submit Request Set").OracleButton("Submit").Click
	wait(2)
	RACK_ReportEvent "Validation Screenshot", "Synchronize Workflow Parameters successfully Enter   ","Screenshot"
	Notification_Message = Get_Oracle_Notification_Form_Message()
	Arr=split(Notification_Message, " ")
	Content = Arr(6)
	Payment_Request_ID_Arr = Split(Content,")")
	Payment_Request_ID = Payment_Request_ID_Arr(0)
	Update_Notepad "Payment_Request_ID", Payment_Request_ID
	Handle_Oracle_Notification_Forms("No")
	wait(2)
	OracleFormWindow("Find Requests").CloseWindow
	Verify_Request_Status(Payment_Request_ID)
End Function
'#####################################################################################################################
'Function Description   : Function to Synchronize Active Users
'Input Parameters 	: None
'Return Value    	: None
'Created By : Avvaru Nagarjuna
'##################################################################################################################### 
Function Active_Users()
   	Select_Link(RACK_GetData("Sys_Admin_Data", "Responsibility_Link"))
	wait(10)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_32").Link("Requests").Click
	wait(60)
	OracleFormWindow("Find Requests").OracleButton("Submit a New Request...").Click
	wait(2)
    OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	wait(2)
	'If  OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Exist(30) Then
		OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Sys_Admin_Data","Request_Name")
		wait(5)
		'OracleFormWindow("Submit Request").OracleTable("Table").SetFocus,"Schedule"
		OracleFormWindow("Submit Request").OracleButton("At these Times...|Schedule...").Click
		wait(2)
		OracleFormWindow("Schedule").OracleRadioGroup("Run the Job..._2").Select "Periodically"
		wait(2)
		OracleFormWindow("Schedule").OracleTextField("End At_2").Enter RACK_GetData("Sys_Admin_Data", "End_Date")
		wait(2)
		OracleFormWindow("Schedule").OracleButton("OK_2").Click
		wait(2)
		OracleFormWindow("Submit Request").OracleButton("Upon Completion...|Options...").Click
		wait(5)
		'OracleFormWindow("Upon Completion...").OracleTable("Table_2").Enter RACK_GetData("Sys_Admin_Data", "Name")
		OracleFormWindow("Upon Completion...").OracleTable("Table_2").EnterField 1,"Name",RACK_GetData("Sys_Admin_Data", "Name")
		wait(2)
		OracleFormWindow("Upon Completion...").OracleButton("OK").Click
	wait(2)
	OracleFormWindow("Submit Request").OracleButton("Submit").Click
	RACK_ReportEvent "Validation Screenshot", "Active Users Parameters successfully Enter   ","Screenshot"
	wait(2)
	Notification_Message = Get_Oracle_Notification_Form_Message()
	Arr=split(Notification_Message, " ")
	Content = Arr(6)
	Payment_Request_ID_Arr = Split(Content,")")
	Payment_Request_ID = Payment_Request_ID_Arr(0)
	Update_Notepad "Payment_Request_ID", Payment_Request_ID
	Handle_Oracle_Notification_Forms("No")
	wait(10)
	OracleFormWindow("Find Requests").CloseWindow
	Verify_Request_Status(Payment_Request_ID)
End Function
'#####################################################################################################################
'Function Description   : Function to Rackspace Audit Request Set
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Rackspace_Audit_Request_Set()
   	Select_Link(RACK_GetData("Sys_Admin_Data", "Responsibility_Link"))
	wait(10)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_32").Link("Requests").Click
	wait(60)
	OracleFormWindow("Find Requests").OracleButton("Submit a New Request...").Click
    wait(2)
	OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
	wait(2)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("Sys_Admin_Data","Request_Name")
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
    OracleFlexWindow("Parameters").Approve
	wait(2)
    OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 2,"Parameters"
    OracleFlexWindow("Parameters").Approve
	wait(2)
	OracleFormWindow("Submit Request Set").OracleButton("Submit").Click
	wait(2)
	RACK_ReportEvent "Validation Screenshot", "Rackspace Audit Request Set Parameters successfully Enter   ","Screenshot"
	Notification_Message = Get_Oracle_Notification_Form_Message()
	Arr=split(Notification_Message, " ")
	Content = Arr(6)
	Payment_Request_ID_Arr = Split(Content,")")
	Payment_Request_ID = Payment_Request_ID_Arr(0)
	Update_Notepad "Payment_Request_ID", Payment_Request_ID
	Handle_Oracle_Notification_Forms("No")
	wait(10)
	OracleFormWindow("Find Requests").CloseWindow
	Verify_Request_Status(Payment_Request_ID)
End Function
'#####################################################################################################################
'Function Description   : Function to New concurrent Manager
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function New_Concurrent_Manager()
   	Select_Link(RACK_GetData("Sys_Admin_Data", "Responsibility_Link"))
	wait(10)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_34").Link("Define").Click
	wait(60)
	OracleFormWindow("Concurrent Managers").OracleTextField("Manager").Enter RACK_GetData("Sys_Admin_Data", "Manager")
	wait(3)
    OracleFormWindow("Concurrent Managers").OracleTextField("Short Name").Enter RACK_GetData("Sys_Admin_Data", "Short_Name")
	wait(3)
    OracleFormWindow("Concurrent Managers").OracleTextField("Application Name").Enter RACK_GetData("Sys_Admin_Data", "Application_Name")
	wait(3)
    OracleFormWindow("Concurrent Managers").OracleTextField("Description").Enter "Rs invetory manager for applicationis created inventor"
	wait(3)
    OracleFormWindow("Concurrent Managers").OracleTextField("Parallel Concurrent Processing").Enter "DAPP2"
    wait(3)
    OracleFormWindow("Concurrent Managers").OracleTextField("Parallel Concurrent Processing_2").Enter "DAPP2"
    wait(3)
    OracleFormWindow("Concurrent Managers").OracleTextField("Program Library|Name").Enter "INVLIBR"
    wait(3)
	OracleFormWindow("Concurrent Managers").SelectMenu "File->Save"
	Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
End Function



'#####################################################################################################################
'Function Description   : Function to Login to EBS Application
'Input Parameters 	: Link Name
'Return Value    	: None
'##################################################################################################################### 

Function Login() 
'		If Browser("Login").Page("Oracle Applications Home_2").Link("Logout").Exist(1) Then
'			 Browser("Login").Page("Oracle Applications Home_2").Link("Logout").Click
'		end if
'        If 	Browser("Login").Exist(1) Then
'			Browser("Login").CloseAllTabs()
'		end if
'		If Browser("Oracle Applications R12").Exist(1) Then
'			Browser("Oracle Applications R12").CloseAllTabs()
'		End If		
	'While Browser("CreationTime:=0").Exist
		'Browser("CreationTime:=0").Close
	'Wend
'	While Browser("CreationTime:=1").Exist
'		Browser("CreationTime:=1").Close
'	Wend

   'SystemUtil.Run "iexplore","https://patch.orapi.rackspace.com/"
	SystemUtil.Run "iexplore",RACK_GetData("Common_Data", "Application_URL" ),"","",3
	If   Browser("Login").Page("Login").WebEdit("usernameField").Exist(120) Then
        Browser("Login").Page("Login").WebEdit("usernameField").Set RACK_GetData("Common_Data", "username")
		RACK_ReportEvent "Login ", "User name "&RACK_GetData("Common_Data", "username" )& " is entered ","Done"
 		Browser("Login").Page("Login").WebEdit("passwordField").SetSecure RACK_GetData("Common_Data", "password" )
		Browser("Login").Page("Login").WebButton("Login").Click
		RACK_ReportEvent "Login ", "Login button is clicked","Done"
		If Browser("Login").Page("Oracle Applications Home_2").Link("Logout").Exist(20) Then
				Reporter.ReportEvent micPass, "Login ", "Successfully logined"
				RACK_ReportEvent "Login ", "Successfully logined","Done"
		Else 
				RACK_ReportEvent  "Login ", "Home Page does not exist", "Fail"
		End If
	Else
			RACK_ReportEvent  "Login ", "Login Failed - Username text box does not exist", "Fail"
	End If
End Function



'#####################################################################################################################
'Function Description   : Function to select a link after logging in to EBS
'Input Parameters 	: Link Name
'Return Value    	: None
'##################################################################################################################### 

Function Select_Link(StrLink_Name)
	If Browser("name:=Oracle Applications Home Page").Page("title:=Oracle Applications Home Page").Link("name:="& StrLink_Name).Exist(20) then
		Browser("name:=Oracle Applications Home Page").Page("title:=Oracle Applications Home Page").Link("name:="& StrLink_Name).Click
		RACK_ReportEvent "Link Clicked", "The link ' " & StrLink_Name & "' is available and clicked","Done"
	else
 		RACK_ReportEvent "Link Clicked", "The link ' " & StrLink_Name & "' is not available and thus not clicked","Fail"
	End if
End Function

'#####################################################################################################################
'Function Description   : Function to handle Oracle Notification Forms
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Handle_Oracle_Notification_Forms(StrButton_To_Press)

	If OracleNotification("title:=.*").Exist(5) then
		If  OracleNotification("title:=.*").OracleButton("label:=" & StrButton_To_Press).Exist(2) Then
			OracleNotification("title:=.*").OracleButton("label:=" & StrButton_To_Press).Click
		End If
		wait(2)
	end if
End Function

'#####################################################################################################################
'Function Description   : Function to extract Oracle Notification Forms messages
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Get_Oracle_Notification_Form_Message()
	Message = "No Message"
	If OracleNotification("title:=.*").Exist(5) then
		Message = OracleNotification("title:=.*").GetROProperty("message")
	end if
	'msgbox Message
	Get_Oracle_Notification_Form_Message = Message
End Function

'#####################################################################################################################
'Function Description   : Function to verify the Oracle message at Status bar
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Verify_Oracle_Status(Expected_Status_Message)

	Actual_Status = OracleStatusLine("OracleStatusLine").GetROProperty("message")
	If (Actual_Status = Expected_Status_Message) Then
			RACK_ReportEvent "Oracle Status Message", "The Oracle Status Message is correctly displayed as '" & Actual_Status & "'.","Pass"
		Else
			RACK_ReportEvent "Oracle Status Message", "The Oracle Status Message is not correctly displayed as Expected. The Expected Value is ' " & Expected_Status_Message & "' and the displayed Value is '" & Actual_Status & "'.","Fail"
		End If

End Function

'#####################################################################################################################
'Function Description   : Function to close all the oracle forms
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Close_all_forms()
	OracleFormWindow("Transactions").SelectMenu "File->Exit Oracle Applications"
	OracleNotification("Caution").Approve
	wait(10)
End Function


'#####################################################################################################################
'Function Description   : Function to Approve Request
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Approve_Request()
	'msgbox RACK_GetData("P2P_Data", "ApprovalData")
	If RACK_GetData("P2P_Data", "ApprovalData") = "Request" Then
		Approval_Number = Read_Notepad("Request_Number")
	elseif RACK_GetData("P2P_Data", "ApprovalData") = "PO" Then
	    wait(100)
		Approval_Number = Read_Notepad("PO_Number")	
	elseif RACK_GetData("P2P_Data", "ApprovalData") = "INVOICE" Then
	    wait(10)
		Approval_Number = Read_Notepad("Invoice_Number")	
	else
		Approval_Number = RACK_GetData("P2P_Data", "ApprovalData")
	End If
	'msgbox Approval_Number
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Workflow")
	Select_Link(RACK_GetData("P2P_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("P2P_Data", "Functionality_Link"))

	If Parent_desc.WebButton("Reassign").Exist(10)Then
'		Parent_desc.WebCheckBox("N25:selected:0").Set "ON"
		Browser("name:=.*").Page("title:=.*").Link("name:=.*"& Approval_Number & ".*").Click
		wait(1)

		'Parent_desc.WebButton("Open").Click
	End If
	If  Parent_desc.WebEdit("NRR0").Exist(10)Then
		Parent_desc.WebEdit("NRR0").Set "Approve"
	End If
	If Parent_desc.WebButton("Approve").Exist(10)Then
		Parent_desc.WebButton("Approve").Click
		wait(5)
		RACK_ReportEvent "Approval", "The user approved - " & Approval_Number,"Pass"
	End If



'	OracleFormWindow("Transactions").SelectMenu "File->Exit Oracle Applications"
'	OracleNotification("Caution").Approve
'	wait(10)
End Function


'#####################################################################################################################
'Function Description   : Function to Reject Request
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Reject_Request()
	'msgbox RACK_GetData("P2P_Data", "ApprovalData")
	If RACK_GetData("P2P_Data", "ApprovalData") = "Request" Then
		Approval_Number = Read_Notepad("Request_Number")
	elseif RACK_GetData("P2P_Data", "ApprovalData") = "PO" Then
	    wait(10)
		Approval_Number = Read_Notepad("PO_Number")	
	elseif RACK_GetData("P2P_Data", "ApprovalData") = "INVOICE" Then
	    wait(10)
		Approval_Number = Read_Notepad("Invoice_Number")	
	else
		Approval_Number = RACK_GetData("P2P_Data", "ApprovalData")
	End If
	'msgbox Approval_Number
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Workflow")
	Select_Link(RACK_GetData("P2P_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("P2P_Data", "Functionality_Link"))

	If Parent_desc.WebButton("Reassign").Exist(10)Then
'		Parent_desc.WebCheckBox("N25:selected:0").Set "ON"
		Browser("name:=.*").Page("title:=.*").Link("name:=.*"& Approval_Number & ".*").Click
		wait(1)

		'Parent_desc.WebButton("Open").Click
	End If
	If  Parent_desc.WebEdit("NRR0").Exist(10)Then
		Parent_desc.WebEdit("NRR0").Set "Rejected"
	End If
	If Parent_desc.WebButton("Reject").Exist(10)Then
		Parent_desc.WebButton("Reject").Click
		wait(5)
		RACK_ReportEvent "Reject", "The user rejected - " & Approval_Number,"Pass"
	End If



'	OracleFormWindow("Transactions").SelectMenu "File->Exit Oracle Applications"
'	OracleNotification("Caution").Approve
'	wait(10)
End Function


'#####################################################################################################################
'Function Description   : Function to Approve & Forward Request
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Approve_Forward_Request()
	If RACK_GetData("P2P_Data", "ApprovalData") = "Request" Then
		Approval_Number = Read_Notepad("Request_Number")
	elseif RACK_GetData("P2P_Data", "ApprovalData") = "PO" Then
	    wait(10)
		Approval_Number = Read_Notepad("PO_Number")	
	elseif RACK_GetData("P2P_Data", "ApprovalData") = "INVOICE" Then
	    wait(10)
		Approval_Number = Read_Notepad("Invoice_Number")	
	else
		Approval_Number = RACK_GetData("P2P_Data", "ApprovalData")
	End If
	'msgbox Approval_Number
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Workflow")
	Select_Link(RACK_GetData("P2P_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("P2P_Data", "Functionality_Link"))

	If Parent_desc.WebButton("Reassign").Exist(10)Then
'		Parent_desc.WebCheckBox("N25:selected:0").Set "ON"
		Browser("name:=.*").Page("title:=.*").Link("name:=.*"& Approval_Number & ".*").Click
		wait(1)
	End If
	If Parent_desc.WebButton("Approve And Forward").Exist(10)Then
		Parent_desc.WebEdit("wfUserName1").click
		WshShell.SendKeys RACK_GetData("P2P_Data", "Forward_To")
		WshShell.SendKeys "{TAB}"
		'Parent_desc.WebEdit("wfUserName1").Set RACK_GetData("P2P_Data", "Forward_To")
		wait(30)
		RACK_ReportEvent "Validation Screenshot", "Confirmation maseage sucessfully displyed  ","Screenshot"
		Parent_desc.WebButton("Approve And Forward").Click
		wait(5)
		RACK_ReportEvent "Approve And Forward", "The user approved and Forwarded - " & Approval_Number,"Pass"
	End If



'	OracleFormWindow("Transactions").SelectMenu "File->Exit Oracle Applications"
'	OracleNotification("Caution").Approve
'	wait(10)
End Function

'#####################################################################################################################
'Function Description   : Function to Update a Notepad File with details
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Update_Notepad(File_Name, Data_To_Update)
	Dim fso, MyFile
	File_path = Environment.Value("ResultPath")  & "\" & File_Name & ".txt"
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set MyFile = fso.CreateTextFile(File_path, True)
	MyFile.WriteLine(Data_To_Update)
	MyFile.Close
End Function

'#####################################################################################################################
'Function Description   : Function to read a Notepad File with details
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Read_Notepad(File_Name)
	Const ForReading = 1
	Set objFS = CreateObject("Scripting.FileSystemObject")
	File_path = Environment.Value("ResultPath")  & "\" & File_Name & ".txt"
	'msgbox File_path
	'If objFS.FileExists(File_path) Then
		Set objFile = objFS.OpenTextFile(File_path, ForReading, False)
		'Do Until objFile.AtEndOfStream
			strLine = objFile.ReadLine
			'MsgBox strLine
		'Loop ' next line
		'objFile.Close
		'Set objFile = Nothing
	'End if ' file exists
	Read_Notepad = strLine
End Function


'#####################################################################################################################
'Function Description   : Function to create todays date in required format
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Format_Date(format)
	MyDate = Date
	MyDay = day(MyDate)
	MyMonth = Monthname(Month(MyDate), true)
	MyYear = Year(MyDate)
	If format = "dd-mmm-yyyy" Then
		Today_Date = MyDay & "-" & MyMonth & "-" & MyYear
	End If
	If format = "dd+1-mmm-yyyy" Then
		MyDay = MyDay + 1
		Today_Date = MyDay & "-" & MyMonth & "-" & MyYear
	End If
	If format = "mmm-yy" Then
		Today_Date = MyMonth & "-" & Right(MyYear, 2)
	End If

	'msgbox Today_Date
	Format_Date = Today_Date
End Function

'#####################################################################################################################
'Function Description   : Function to close all Oracle Windows
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Close_Oracle_forms()
	OracleFormWindow("Navigator").CloseWindow
	'OracleFormWindow("Navigator").CloseForm
	Handle_Oracle_Notification_Forms("Discard")
	Handle_Oracle_Notification_Forms("OK")
	wait(5)
End Function

'#####################################################################################################################
'Function Description   : Function to Verify Request Status and Phase
'Input Parameters 	: None
'Return Value    	: None
'#####################################################################################################################
Function Verify_Request_Status(Request_ID)

	OracleFormWindow("Navigator").SelectMenu "View->Requests"
	wait(5)
	If OracleFormWindow("Find Requests").OracleRadioGroup("My Completed Requests").Exist Then
		OracleFormWindow("Find Requests").OracleRadioGroup("My Completed Requests").Select "Specific Requests"
	else if  OracleFormWindow("Find Requests").OracleRadioGroup("My Completed Requests_2").Exist Then
		OracleFormWindow("Find Requests").OracleRadioGroup("My Completed Requests_2").Select "Specific Requests"
	End if
	End If
	OracleFormWindow("Find Requests").OracleTextField("Request ID").Enter Request_ID
	'wait(30)
	OracleFormWindow("Find Requests").OracleButton("Find").Click
	If OracleFormWindow("Requests").Exist(20) Then
		'wait(30)
		phase_Value = OracleFormWindow("Requests").OracleTable("Table").GetFieldValue(1,"Phase")
		'phase_Value = OracleFormWindow("Requests").OracleTextField("Phase").GetROProperty("value")
		'msgbox phase_Value
		While not phase_Value = "Completed"
			wait(3)
			OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			wait(3)
			phase_Value = OracleFormWindow("Requests").OracleTable("Table").GetFieldValue( 1,"Phase")
			'phase_Value = OracleFormWindow("Requests").OracleTextField("Phase").GetROProperty("value")
		Wend
		'msgbox "Pass"
		'OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			If  OracleFormWindow("Requests").OracleTable("Table").GetFieldValue(1,"Phase") = "Completed" Then
		 '1,"Phase" = "Completed") Then
		'If (OracleFormWindow("Requests").OracleTextField("Phase").GetROProperty("value") = "Completed") Then
			RACK_ReportEvent "Request Phase", "The Request Phase is correctly displayed as '" & OracleFormWindow("Requests").OracleTable("Table").GetFieldValue(1,"Phase")& "'.","Pass"
			RACK_ReportEvent "Validation Screenshot", "Validation Screenshot" ,"Screenshot"
		Else
			RACK_ReportEvent "Request Phase", "The Request Phase is not correctly displayed as Expected. The Expected Value is 'Completed' and displayed Value is '" & OracleFormWindow("Requests").OracleTable("Table").GetFieldValue(1,"Phase")& "'.","Fail"
		End If
		If (OracleFormWindow("Requests").OracleTable("Table").GetFieldValue(1,"Status") = "Normal") Then
			RACK_ReportEvent "Request Status", "The Request Status is correctly displayed as '" & OracleFormWindow("Requests").OracleTable("Table").GetFieldValue(1,"Status")& "'.","Pass"
			RACK_ReportEvent "Validation Screenshot", "Validation Screenshot" ,"Screenshot"
		Else
			RACK_ReportEvent "Request Status", "The Request Status is not correctly displayed as Expected. The Expected Value is 'Normal' and displayed Value is '" & OracleFormWindow("Requests").OracleTable("Table").GetFieldValue(1,"Status")& "'.","Fail"
		End If
   	End if
End Function


'#####################################################################################################################
'Function Description   : Function to Switch the Responsibility
'Input Parameters 	: None
'Return Value    	: None
'#####################################################################################################################
Function Switch_Responsibility(Responsibility)
	wait(3)
	OracleFormWindow("Navigator").SelectMenu "File->Switch Responsibility..."
	OracleListOfValues("Responsibilities").Select Responsibility

End Function



'#####################################################################################################################
'Function Description   : Function to Login to Portal
'Input Parameters 	: Link Name
'Return Value    	: None
'##################################################################################################################### 

Function Login_Portal() 
	SystemUtil.Run "iexplore",RACK_GetData("Common_Data", "Application_URL" ),"","",3
	wait(20)
	If  Browser("Oracle Applications Home Page").Page("Log In — MyRackspace").WebEdit("username").Exist(120) Then
		Browser("Oracle Applications Home Page").Page("Log In — MyRackspace").WebEdit("account").Set RACK_GetData("Common_Data", "Account_Number")
        Browser("Oracle Applications Home Page").Page("Log In — MyRackspace").WebEdit("username").Set RACK_GetData("Common_Data", "username")
		RACK_ReportEvent "Login ", "User name "&RACK_GetData("Common_Data", "username" )& " is entered ","Done"
 		Browser("Oracle Applications Home Page").Page("Log In — MyRackspace").WebEdit("password").SetSecure RACK_GetData("Common_Data", "password" )
		wait(2)
        RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		Browser("Oracle Applications Home Page").Page("Log In — MyRackspace").WebButton("Log In").Click
		RACK_ReportEvent "Login ", "Login button is clicked","Done"
'		If Browser("Oracle Applications Home Page").Page("Log In — MyRackspace").Link("Log Out").Exist(10) Then
'				Reporter.ReportEvent micPass, "Login ", "Successfully logined"
'				RACK_ReportEvent "Login ", "Successfully logined","Done"
'		Else 
'				RACK_ReportEvent  "Login ", "Home Page does not exist", "Fail"
'		End If
'	Else
'			RACK_ReportEvent  "Login ", "Login Failed - Username text box does not exist", "Fail"
	End If
End Function



'#####################################################################################################################
'Function Description   : Function to Login to Portal
'Input Parameters 	: Link Name
'Return Value    	: None
'##################################################################################################################### 

Function Logout_Portal() 
	If  Browser("Oracle Applications Home Page").Page("Log In — MyRackspace").Link("Log Out").Exist(5) Then
		Browser("Oracle Applications Home Page").Page("Log In — MyRackspace").Link("Log Out").Click
		wait(2)
		SystemUtil.CloseProcessByName("IExplore.exe")
	End If
End Function

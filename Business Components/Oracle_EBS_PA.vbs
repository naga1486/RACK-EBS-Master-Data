
'#####################################################################################################################
'Function Description   : Function to Verify Project Expenditure
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_New_Project_Organization()
	Select_Link(RACK_GetData("PA_Data", "Responsibility_Link"))
	wait(5)
	'Select_Link(RACK_GetData("PA_Data", "Functionality_Link"))
	 Browser("Oracle Applications Home Page").Page("Oracle Applications Home_40").Link("Description").Click
	If OracleFormWindow("Find Organization").Exist(120) Then
	OracleFormWindow("Find Organization").OracleButton("New").Click
	wait(2)
    OracleFormWindow("Organization").OracleTextField("Name").Enter RACK_GetData("PA_Data", "Organization")
    wait(2)
    OracleFormWindow("Organization").SelectMenu "File->Save"
	Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
	wait(2)
    OracleFormWindow("Organization").OracleTable("Table").EnterField 1,"Name","Project Expenditure/Event Organization"
    OracleFormWindow("Organization").OracleTable("Table").EnterField 1,"Enabled",true
	wait(2)
    OracleFormWindow("Organization").SelectMenu "File->Save"
	wait(2)
    OracleFormWindow("Organization").OracleTable("Table").EnterField 2,"Name","Project Task Owning Organization"
    OracleFormWindow("Organization").OracleTable("Table").EnterField 2,"Enabled",true
	wait(2)
    OracleFormWindow("Organization").SelectMenu "File->Save"
	wait(2)
    OracleFormWindow("Organization").OracleTable("Table").EnterField 3,"Name","Project Invoice Collection Organization"
    OracleFormWindow("Organization").OracleTable("Table").EnterField 3,"Enabled",true
	wait(2)
    OracleFormWindow("Organization").SelectMenu "File->Save"
	Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
	wait(2)
	End If
End Function
'#####################################################################################################################
'Function Description   : Function to Create a new Project
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Global_Organization_Hierachy()
	Select_Link(RACK_GetData("PA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("PA_Data", "Functionality_Link"))
	If OracleFormWindow("Global Organization Hierarchy").Exist(120) Then
	OracleFormWindow("Global Organization Hierarchy").SelectMenu "View->Query By Example->Enter"
	wait(2)
    OracleFormWindow("Global Organization Hierarchy").SelectMenu "View->Query By Example->Run"
	wait(2)
    'OracleFormWindow("Global Organization Hierarchy").OracleTextField("Organization|Name").SetFocus
    OracleFormWindow("Global Organization Hierarchy").OracleTextField("Organization|Name").SetFocus
	wait(2)
    OracleFormWindow("Global Organization Hierarchy").SelectMenu "View->Query By Example->Enter"
	wait(2)
    OracleFormWindow("Global Organization Hierarchy").OracleTextField("Organization|Name").Enter "Rackspace US - Operating Unit"
	wait(2)
    OracleFormWindow("Global Organization Hierarchy").SelectMenu "View->Query By Example->Run"
	wait(5)
    'OracleFormWindow("Global Organization Hierarchy").OracleTable("Table").SetFocus
	'OracleFormWindow("Global Organization Hierarchy").OracleTable("Table").SetFocus
	OracleFormWindow("Global Organization Hierarchy").OracleTable("Table").SetFocus 1,"Name"
    OracleFormWindow("Global Organization Hierarchy").SelectMenu "File->New"
	wait(2)
    OracleFormWindow("Global Organization Hierarchy").OracleTable("Table").EnterField 2,"Name",RACK_GetData("PA_Data", "Organization")
	wait(2)
    OracleFormWindow("Global Organization Hierarchy").SelectMenu "File->Save"
	Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
	wait(2)
	End If
End Function
'#####################################################################################################################
'Function Description   : Function to Assign  Project Lookup Sets
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Lookup_Sets()
	Select_Link(RACK_GetData("PA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("PA_Data", "Functionality_Link"))
	If OracleFormWindow("AutoAccounting Lookup").Exist(120) Then
        OracleFormWindow("AutoAccounting Lookup").SelectMenu "View->Query By Example->Enter"
		wait(2)
        OracleFormWindow("AutoAccounting Lookup").OracleTextField("Name").Enter "Dept. to Company Account"
		wait(2)
        OracleFormWindow("AutoAccounting Lookup").SelectMenu "View->Query By Example->Run"
		OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").SetFocus 1,"Intermediate Value"
		wait(2)
        OracleFormWindow("AutoAccounting Lookup").SelectMenu "File->New"
		wait(2)
        OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").EnterField 2,"Intermediate Value",RACK_GetData("PA_Data", "Organization")
		wait(2)
        OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").EnterField 2,"Segment Value","100"
		wait(2)
        OracleFormWindow("AutoAccounting Lookup").SelectMenu "File->Save"
        Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
		wait(2)
         OracleFormWindow("AutoAccounting Lookup").OracleTextField("Name").SetFocus
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").SelectMenu "View->Query By Example->Enter"
		wait(2)
        OracleFormWindow("AutoAccounting Lookup").OracleTextField("Name").Enter "Dept. to Business Unit Account"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").SelectMenu "View->Query By Example->Run"
		OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").SetFocus 1,"Intermediate Value"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").SelectMenu "File->New"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").EnterField 2,"Intermediate Value",RACK_GetData("PA_Data", "Organization")
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").EnterField 2,"Segment Value","5400"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").SelectMenu "File->Save"
        Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
		wait(2)
		 OracleFormWindow("AutoAccounting Lookup").OracleTextField("Name").SetFocus
		wait(2)
        OracleFormWindow("AutoAccounting Lookup").SelectMenu "View->Query By Example->Enter"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTextField("Name").Enter "Dept. to Team Account"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").SelectMenu "View->Query By Example->Run"
		OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").SetFocus 1,"Intermediate Value"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").SelectMenu "File->New"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").EnterField 2,"Intermediate Value",RACK_GetData("PA_Data", "Organization")
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").EnterField 2,"Segment Value","000"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").SelectMenu "File->Save"
        Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTextField("Name").SetFocus
		wait(2)
        OracleFormWindow("AutoAccounting Lookup").SelectMenu "View->Query By Example->Enter"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTextField("Name").Enter "Exp. Type to Natural Account"
		wait(2)
        OracleFormWindow("AutoAccounting Lookup").SelectMenu "View->Query By Example->Run"
		OracleFormWindow("AutoAccounting Lookup").SelectMenu "File->Save"
        RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(2)
		'OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").SetFocus 1,"Intermediate Value"
		'wait(2)
        'OracleFormWindow("AutoAccounting Lookup").SelectMenu "File->New"
		'wait(2)
       ' OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").EnterField 2,"Intermediate Value",RACK_GetData("PA_Data", "Organization")
		'wait(2)
        'OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").EnterField 2,"Segment Value","000"
		OracleFormWindow("AutoAccounting Lookup").OracleTextField("Name").SetFocus
		wait(2)
        OracleFormWindow("AutoAccounting Lookup").SelectMenu "View->Query By Example->Enter"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTextField("Name").Enter "Location to Account Segment"
		wait(2)
        OracleFormWindow("AutoAccounting Lookup").SelectMenu "View->Query By Example->Run"
		OracleFormWindow("AutoAccounting Lookup").SelectMenu "File->Save"
        RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTextField("Name").SetFocus
		wait(2)
		 OracleFormWindow("AutoAccounting Lookup").SelectMenu "View->Query By Example->Enter"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTextField("Name").Enter "Project Organization to Dept."
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").SelectMenu "View->Query By Example->Run"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").SetFocus 1,"Intermediate Value"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").EnterField 2,"Intermediate Value",RACK_GetData("PA_Data", "Organization")
		wait(5)
		OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").EnterField 2,"Segment Value","3107"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").SelectMenu "File->Save"
        Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTextField("Name").SetFocus
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").SelectMenu "View->Query By Example->Enter"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTextField("Name").Enter "Project Org to Product"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").SelectMenu "View->Query By Example->Run"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").SetFocus 1,"Intermediate Value"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").EnterField 2,"Intermediate Value",RACK_GetData("PA_Data", "Organization")
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").OracleTable("Table").EnterField 2,"Segment Value","0000"
		wait(2)
		OracleFormWindow("AutoAccounting Lookup").SelectMenu "File->Save"
        Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
		wait(2)
		End If
End Function
'#####################################################################################################################
''Function Description   : Function to split expenditure lines
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 
Function Create_Project_With_New_Org()
'set WshShell = CreaCteObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("PA_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("PA_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Find Projects").OracleTextField("Project|Number").Enter RACK_GetData("PA_Data", "Project_Number")
	wait(2)
    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	OracleFormWindow("Find Projects").OracleButton("Find").Click
'	wait(2)
'	OracleFormWindow("Projects, Templates Summary").OracleButton("Open").Click
	wait(2)
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	OracleFormWindow("Projects, Templates Summary").OracleButton("Copy To...").Click
	wait(2)
	OracleFormWindow("Project Quick Entry").OracleTable("Table").EnterField 1,"Value",RACK_GetData("PA_Data", "Project_Name")
	wait(2)
    OracleFormWindow("Project Quick Entry").OracleTable("Table").EnterField 2,"Value",RACK_GetData("PA_Data", "Project_Description")
	wait(2)
    OracleFormWindow("Project Quick Entry").OracleTable("Table").EnterField 3,"Value",RACK_GetData("PA_Data", "Project_Manager")
	wait(2)
	'OracleFormWindow("Project Quick Entry").OracleTable("Table").EnterField 4,"Value", Format_Date("dd-mmm-yyyy")
	OracleFormWindow("Project Quick Entry").OracleTable("Table").EnterField 4,"Value",RACK_GetData("PA_Data", "Project_Start_Date")
	wait(2)
	OracleFormWindow("Project Quick Entry").OracleTable("Table").EnterField 5,"Value",RACK_GetData("PA_Data", "Project_Organization")
	wait(2)
	OracleFormWindow("Project Quick Entry").OracleTable("Table").EnterField 6,"Value",RACK_GetData("PA_Data", "Project_Location")
	wait(2)
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	OracleFormWindow("Project Quick Entry").OracleButton("OK").Click
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(5)
	'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	OracleFormWindow("Projects, Templates Summary").OracleButton("Open").Click
    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	OracleFormWindow("Projects, Templates").OracleButton("Change Status").Click
	'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
    OracleListOfValues("Project Status").Select "Approved"
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	OracleFormWindow("Projects, Templates").OracleButton("Detail").Click
	wait(5)
    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	'OracleFormWindow("Tasks").SelectMenu "File->New"
	'wait(2)
	'OracleFormWindow("Tasks").OracleTable("Table").EnterField 2,"Task Number",RACK_GetData("PA_Data", "Task_Number")
	'wait(2)
    'OracleFormWindow("Tasks").OracleTable("Table").EnterField 2,"Task Name",RACK_GetData("PA_Data", "Task_Name")
	'ait(2)
    'acleFormWindow("Tasks").OracleTable("Table").EnterField 2,"Description",RACK_GetData("PA_Data", "Project_Description")
	'wait(2)
    OracleFormWindow("Tasks").CloseWindow
    OracleFormWindow("Projects, Templates").CloseWindow
	wait(5)
    Verify_Oracle_Status("FRM-40400: Transaction complete: 2 records applied and saved.")
End Function

'#####################################################################################################################
'Function Description   : Function to Create Purchase Order
'Input Parameters 	: None
'Return Value    	: None
'#####################################################################################################################

Function Create_PO_With_New_Project()
	Select_Link(RACK_GetData("PA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("PA_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Purchase Orders").OracleTextField("Type").SetFocus
	If not  RACK_GetData("PA_Data", "PO_Type") = "" Then
		OracleFormWindow("Purchase Orders").OracleTextField("Type").Enter RACK_GetData("PA_Data", "PO_Type")
	End If
	OracleFormWindow("Purchase Orders").OracleTextField("Supplier").Enter RACK_GetData("PA_Data", "Supplier_Name")
	If not RACK_GetData("PA_Data", "Supplier_To_Site") = "" Then
		OracleFormWindow("Purchase Orders").OracleTextField("Site").Enter RACK_GetData("PA_Data", "Supplier_To_Site")
	End If
	If not RACK_GetData("PA_Data", "PO_Item") = "" Then
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").SetFocus 1,"Item"
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Item",RACK_GetData("PA_Data", "PO_Item")
	End If
	If not RACK_GetData("PA_Data", "PO_Category") = "" Then
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").SetFocus 1,"Category"
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Category",RACK_GetData("PA_Data", "PO_Category")
			wait(2)
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Description",RACK_GetData("PA_Data", "PO_Description")
	End If
	OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Quantity",RACK_GetData("PA_Data", "PO_Quantity")
	wait(2)
	OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table_3").EnterField 1,"Price",RACK_GetData("PA_Data", "PO_Price")
	wait(2)
	OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table_4").EnterField 1,"Need-By",RACK_GetData("PA_Data", "PO_Need_By_Date")
	wait(2)
	RACK_ReportEvent "Validation Screenshot", "Purchase Orders Parameters successfully Enter  ","Screenshot"
	OracleFormWindow("Purchase Orders").OracleButton("Shipments").Click
	wait(2)
	OracleFormWindow("Shipments").OracleButton("Distributions").Click
	wait(2)
    OracleFormWindow("Distributions").OracleTabbedRegion("Project").OracleTable("Table").EnterField 1,"Project",RACK_GetData("PA_Data", "Project_Number")
	wait(2)
	OracleFormWindow("Distributions").OracleTabbedRegion("Project").OracleTable("Table").EnterField 1,"Task",RACK_GetData("PA_Data", "Task_Number")
	wait(2)
	OracleFormWindow("Distributions").OracleTabbedRegion("Project").OracleTable("Table").EnterField 1,"Type",RACK_GetData("PA_Data", "Type")
	 wait(2)
	 OracleFormWindow("Distributions").OracleTabbedRegion("Project").OracleTable("Table").EnterField 1,"Org",RACK_GetData("PA_Data", "Organization")
     wait(2)
	OracleFormWindow("Distributions").OracleTabbedRegion("Project").OracleTable("Table").EnterField 1,"Date",RACK_GetData("PA_Data", "Date")
    wait(2)
	OracleFormWindow("Distributions").OracleTabbedRegion("Destination").OracleTable("Table").EnterField 1,"Quantity","5"
	 wait(2)
	 'OracleFormWindow("Distributions").OracleTabbedRegion("Table").SetFocus 2,"Intermediate Value"
	OracleFormWindow("Distributions").OracleTabbedRegion("Project").OracleTable("Table").EnterField 2,"Project",RACK_GetData("PA_Data", "Project_Number")
	wait(2)
	OracleFormWindow("Distributions").OracleTabbedRegion("Project").OracleTable("Table").EnterField 2,"Task",RACK_GetData("PA_Data", "Task_Number")
	wait(2)
	OracleFormWindow("Distributions").OracleTabbedRegion("Project").OracleTable("Table").EnterField 2,"Type",RACK_GetData("PA_Data", "Type")
	 wait(2)
	 OracleFormWindow("Distributions").OracleTabbedRegion("Project").OracleTable("Table").EnterField 2,"Org",RACK_GetData("PA_Data", "Organization")
     wait(2)
	OracleFormWindow("Distributions").OracleTabbedRegion("Project").OracleTable("Table").EnterField 2,"Date",RACK_GetData("PA_Data", "Date")
    wait(2)
	OracleFormWindow("Distributions").OracleTabbedRegion("Destination").OracleTable("Table").EnterField 2,"Quantity","5"
	wait(2)
	OracleFormWindow("Distributions").SelectMenu "File->Save"
	wait(5)
	Verify_Oracle_Status("FRM-40400: Transaction complete: 5 records applied and saved.")
    wait(2)
	OracleFormWindow("Distributions").CloseWindow
	wait(2)
	OracleFormWindow("Shipments").CloseWindow
	wait(2)
	Purchase_Order_Number = OracleFormWindow("Purchase Orders").OracleTextField("PO, Rev").GetROProperty("value")
	RACK_PutData "PA_Data", "Purchase_Order_Number", CStr(Purchase_Order_Number)
	Update_Notepad "PO_Number", Purchase_Order_Number
	If  not Purchase_Order_Number = "" Then
	RACK_ReportEvent "Purchase Order", "The Purchase Order '" & Purchase_Order_Number & "' is created sucessfully as expected","Pass"
	else
	RACK_ReportEvent "Purchase Order", "The Purchase Order is not created sucessfully","Fail"
	end if
	OracleFormWindow("Purchase Orders").OracleButton("Approve...").Click
	If  OracleFormWindow("Approve Document").Exist(10) Then
		OracleFormWindow("Approve Document").OracleTabbedRegion("Approval Details").OracleCheckbox("Approval|Submit for Approval").Select
		wait(2)
		OracleFormWindow("Approve Document").OracleTabbedRegion("Approval Details").OracleCheckbox("Approval|Forward").Select
		RACK_ReportEvent "Validation Screenshot", "Approve Document Parameters successfully Enter  ","Screenshot"
		OracleFormWindow("Approve Document").OracleButton("OK").Click
		wait(2)
	end if	
End Function

'#####################################################################################################################
'Function Description   : Function to create a Invoice Batches
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_AP_Invoice()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("PA_Data", "Responsibility_Link"))
	wait(5)
	Parent_desc.Link("Invoice Batches").click
		If OracleFormWindow("Invoice Batches").OracleTable("Table").Exist(120) Then
            OracleFormWindow("Invoice Batches").OracleTable("Table").EnterField 1,"Batch Name", RACK_GetData("PA_Data", "BatchName")
			RACK_ReportEvent "Validation Screenshot", "Po BatchName successfully Enter ","Screenshot"
			wait(2)
			OracleFormWindow("Invoice Batches").OracleButton("Invoices").Click
			wait(5)
			If OracleFormWindow("Invoice Workbench").OracleTable("Table").Exist(20) Then
				If not RACK_GetData("PA_Data", "Purchase_Order_Number") = "" Then
					PO_Number = RACK_GetData("PA_Data", "Purchase_Order_Number")
				else
					PO_Number = Read_Notepad("PO_Number")
				End if
                OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"PO Number", PO_Number
				wait(2)
				OracleFormWindow("Invoice Workbench").OracleTable("Table").SetFocus 1,"Invoice Date"                     
				OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Date",RACK_GetData("PA_Data", "Invoice_Date")
				wait(2)    
				If not RACK_GetData("PA_Data", "Supplier_Name") = "" Then
				OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Trading Partner", RACK_GetData("PA_Data", "Supplier_Name")   
				End if
				OracleFormWindow("Invoice Workbench").OracleTable("Table").SetFocus 1,"Invoice Num" 
				wait(5)
				Handle_Oracle_Notification_Forms("Cancel")
				OracleFormWindow("Invoice Workbench").OracleTable("Table").SetFocus 1,"Invoice Num" 
				Invoice_Number = RACK_GetData("PA_Data", "Invoice_Number")
				OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Num", Invoice_Number
				Update_Notepad "Invoice_Number", Invoice_Number
				OracleFormWindow("Invoice Workbench").OracleTable("Table").SetFocus 1,"Invoice Amount"
				OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Amount", RACK_GetData("PA_Data", "Invoice_Amount")
'				RACK_ReportEvent "Validation Screenshot", "Invoice Workbench Parameters successfully Enter   ","Screenshot"
				wait(2)
				If not RACK_GetData("PA_Data", "Requester") = "" Then
				OracleFormWindow("Invoice Workbench").SelectMenu "Folder->Folder Tools"
				wait(2)
                OracleFormWindow("Folder Tools").OracleButton("Show Field...").Click
				wait(2)
                OracleListOfValues("Show Field").Select "Requester                                                                                                                                                                                                                  *REQUESTER_NAME"
                OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 1,"Requester",RACK_GetData("PA_Data", "Requester")
				wait(2)
                OracleFormWindow("Folder Tools").CloseWindow
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				End If
				wait(2)
				OracleFormWindow("Invoice Workbench").OracleButton("Match").Click
				RACK_ReportEvent "Validation Screenshot", "Po Match Button successfully click ","Screenshot"
				If  OracleFormWindow("Find Purchase Orders for").Exist(10) Then
					OracleFormWindow("Find Purchase Orders for").OracleButton("Find").Click
				End if
				'wait(5)
				If  OracleFormWindow("Match to Purchase Orders").Exist(5) Then
                    OracleFormWindow("Match to Purchase Orders").OracleTable("Table").EnterField 1,"Match", true
				RACK_ReportEvent "Validation Screenshot", "Po Match checkbox successfully select ","Screenshot"
				wait(2)
				OracleFormWindow("Match to Purchase Orders").OracleButton("Match").Click
				wait(5)
                Verify_Oracle_Status("FRM-40400:Transaction complete: 1 records applied and saved.")
				End If
				End If
				If  RACK_GetData("PA_Data", "Invoice_Actions") = "Yes" Then
				     OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
				wait(5)
				OracleFormWindow("Invoice Actions").OracleCheckbox("Validate").Select
				RACK_ReportEvent "Validation Screenshot", "Sucessfully Validate check box Select  ","Screenshot"
				wait(2)
				OracleFormWindow("Invoice Actions").OracleButton("OK").Click
				wait(2)
				OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
				wait(2)
                OracleFormWindow("Invoice Actions").OracleCheckbox("Create Accounting").Select
				OracleFormWindow("Invoice Actions").OracleRadioGroup("Draft").Select "Final"
				wait(2)
				OracleFormWindow("Invoice Actions").OracleCheckbox("Initiate Approval").Select
				RACK_ReportEvent "Validation Screenshot", "Sucessfully Create Accounting check box Select  ","Screenshot"
				wait(2)
                OracleFormWindow("Invoice Actions").OracleButton("OK").Click
				wait(2)
                Verify_Oracle_Status("FRM-40400:Transaction complete: 1 records applied and saved.")
			   End If
			If OracleNotification("Note").Exist(20) Then
				OracleNotification("Note").Approve
	     		Verify_Oracle_Status("FRM-40400:Transaction complete: 1 records applied and saved.")
End If
End If
End Function
'#####################################################################################################################
'Function Description   : Function to Create Accounting
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Interface_Invoice_Into_PA()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("PA_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("PA_Data", "Functionality_Link"))
    If OracleFormWindow("Submit a New Request").Exist(90) Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(10) Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("PA_Data", "Request_Name")
			wait(2)
            OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Parameters").SetFocus
            OracleFlexWindow("Parameters").OracleTextField("Project Number").Enter RACK_GetData("PA_Data", "Project_Number")
            OracleFlexWindow("Parameters").Approve
			'End If
			wait(2)
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
			RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			wait(5)
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

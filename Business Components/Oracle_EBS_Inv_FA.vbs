
'#####################################################################################################################
'Function Description   : Function to create a Item
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_New_Item()
	'set WshShell = CreateObject("WScript.Shell")
	'set Parent_desc = Browser("Oracle Applications Home Page").Page("Suppliers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If OracleFormWindow("Master Item").Exist(100) Then
		OracleFormWindow("Master Item").OracleTextField("Item").Enter RACK_GetData("Inv_FA_Data", "Item_Number")
		wait(2)
		OracleFormWindow("Master Item").OracleTextField("Description").Enter RACK_GetData("Inv_FA_Data", "Item_Description")
		wait(2)
		OracleFormWindow("Master Item").SelectMenu "Tools->Copy From..."
		wait(2)
				If OracleFormWindow("Copy From").OracleTextField("Template").Exist(20) Then
					OracleFormWindow("Copy From").OracleTextField("Template").Enter RACK_GetData("Inv_FA_Data", "Item_Template")
					wait(2)
					RACK_ReportEvent "Validation Screenshot", "Item_Template successfully Enter   ","Screenshot"
					OracleFormWindow("Copy From").OracleButton("Apply").Click
					wait(2)
					OracleFormWindow("Copy From").OracleButton("Done").Click
				End if
		wait(2)
        OracleFormWindow("Master Item").SelectMenu "File->Save"
		wait(2)
		Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
		'RACK_ReportEvent "Item Creation", "The Item '" & RACK_GetData("Inv_FA_Data", "Item_Name") & "'  has been sucessfully created" ,"Pass "
  End if 
End Function

'#####################################################################################################################
'Function Description   : Function to update a Item
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Update_Item()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Suppliers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If OracleFormWindow("Master Item").Exist(100) Then
		OracleFormWindow("Master Item").OracleTextField("Item").InvokeSoftkey "ENTER QUERY"
		wait(2)
		OracleFormWindow("Master Item").OracleTextField("Item").Enter RACK_GetData("Inv_FA_Data", "Item_Number")
		wait(2)
		OracleFormWindow("Master Item").OracleTextField("Item").InvokeSoftkey "EXECUTE QUERY"
		wait(2)
		OracleFormWindow("Master Item").OracleTabbedRegion("Main").OracleTextField("User Item Type").Enter RACK_GetData("Inv_FA_Data", "Item_Template")
		wait(2)
        OracleFormWindow("Master Item").SelectMenu "Tools->Categories"
        RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		If OracleFormWindow("Category Assignment").OracleTable("Table").Exist(20) Then
		'If OracleFormWindow("Category Assignment").OracleTextField("Category").Exist(20) Then
			'OracleFormWindow("Category Assignment").OracleTextField("Category").OpenDialog
          '  OracleFormWindow("Category Assignmen").OracleTable("Table").OpenDialog 1,"Category"
			'OracleListOfValues("RS Item Categories").Select RACK_GetData("Inv_FA_Data", "Item_Category")
			OracleFormWindow("Category Assignment").OracleTable("Table").SetFocus 1, "Category"
			OracleFormWindow("Category Assignment").OracleTable("Table").EnterField 1, "Category",RACK_GetData("Inv_FA_Data", "Item_Category")
			wait(2)
			OracleFormWindow("Category Assignment").PressToolbarButton "Clear Record"
			wait(2)
			OracleFormWindow("Category Assignment").SelectMenu "File->Save"
			wait(5)
			'RACK_ReportEvent "Item Update", "The Item '" & RACK_GetData("Inv_FA_Data", "Item_Name") & "'  has been sucessfully updated" ,"Pass"
			Verify_Oracle_Status("FRM-40400: Transaction complete: 2 records applied and saved.")
		End if
	End if 
End Function

'#####################################################################################################################
'Function Description   : Function for Item Organization Assignment
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Item_Organization_Assignment()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Suppliers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If OracleFormWindow("Master Item").Exist(100) Then
		OracleFormWindow("Master Item").OracleTextField("Item").InvokeSoftkey "ENTER QUERY"
		wait(2)
		OracleFormWindow("Master Item").OracleTextField("Item").Enter RACK_GetData("Inv_FA_Data", "Item_Number")
		wait(2)
		OracleFormWindow("Master Item").OracleTextField("Item").InvokeSoftkey "EXECUTE QUERY"
		wait(2)
        OracleFormWindow("Master Item").SelectMenu "Tools->Organization Assignment"
		wait(2)
		'OracleFormWindow("Master Item").OracleCheckbox("Assign Item to Organization").Select
		OracleFormWindow("Master Item").OracleTable("Table").EnterField 8,"Assigned",True
		wait(2)
		OracleFormWindow("Master Item").SelectMenu "File->Save"
        Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
		wait(2)
		OracleFormWindow("Master Item").OracleTable("Table").EnterField 19,"Assigned",True
		wait(1)
		OracleFormWindow("Master Item").OracleTable("Table").EnterField 20,"Assigned",True
		wait(1)
		OracleFormWindow("Master Item").OracleTable("Table").EnterField 21,"Assigned",True
		wait(1)
		OracleFormWindow("Master Item").OracleTable("Table").EnterField 8,"Assigned",True
		wait(2)
		OracleFormWindow("Master Item").SelectMenu "File->Save"
		wait(5)
		'RACK_ReportEvent "Item Organization Assignment", "The Item '" & RACK_GetData("Inv_FA_Data", "Item_Name") & "'  Organization Assignment has been sucessfully done" ,"Pass"
		Verify_Oracle_Status("FRM-40400: Transaction complete: 3 records applied and saved.")
		'OracleFormWindow("Master Item").CloseWindow
	End if
End Function
'#####################################################################################################################
'Function Description   : Function to Create Purchase Order
'Input Parameters 	: None
'Return Value    	: None
'#####################################################################################################################

Function Create_Po()
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Purchase Orders").OracleTextField("Type").SetFocus
	If not  RACK_GetData("Inv_FA_Data", "PO_Type") = "" Then
		OracleFormWindow("Purchase Orders").OracleTextField("Type").Enter RACK_GetData("Inv_FA_Data", "PO_Type")
	End If
	OracleFormWindow("Purchase Orders").OracleTextField("Supplier").Enter RACK_GetData("Inv_FA_Data", "Supplier_Name")
	If not RACK_GetData("Inv_FA_Data", "Supplier_To_Site") = "" Then
		OracleFormWindow("Purchase Orders").OracleTextField("Site").Enter RACK_GetData("Inv_FA_Data", "Supplier_To_Site")
	End If
	If not RACK_GetData("Inv_FA_Data", "PO_Item") = "" Then
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").SetFocus 1,"Item"
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Item",RACK_GetData("Inv_FA_Data", "PO_Item")
	End If
	If not RACK_GetData("Inv_FA_Data", "PO_Category") = "" Then
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").SetFocus 1,"Category"
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Category",RACK_GetData("Inv_FA_Data", "PO_Category")
			wait(2)
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Description",RACK_GetData("Inv_FA_Data", "PO_Description")
	End If
	OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Quantity",RACK_GetData("Inv_FA_Data", "PO_Quantity")
	wait(2)
	OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table_3").EnterField 1,"Price",RACK_GetData("Inv_FA_Data", "PO_Price")
	wait(2)
	OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table_4").EnterField 1,"Need-By",RACK_GetData("Inv_FA_Data", "PO_Need_By_Date")
	wait(2)
	RACK_ReportEvent "Validation Screenshot", "Purchase Orders Parameters successfully Enter  ","Screenshot"
	OracleFormWindow("Purchase Orders").OracleButton("Shipments").Click
	wait(2)
	RACK_ReportEvent "Validation Screenshot", "Shipments Parameters successfully Enter  ","Screenshot"
	OracleFormWindow("Shipments").OracleButton("Distributions").Click
	wait(2)
	If not RACK_GetData("Inv_FA_Data", "Po_Charge_Account") = "" Then
		OracleFormWindow("Distributions").OracleTabbedRegion("Destination").OracleTable("Table").SetFocus 1,"PO Charge Account"
		'OracleFlexWindow("Charge Account").Cancel
		OracleFormWindow("Distributions").OracleTabbedRegion("Destination").OracleTable("Table_2").EnterField 1,"PO Charge Account",RACK_GetData("Inv_FA_Data", "Po_Charge_Account")
		'RACK_ReportEvent "Validation Screenshot", "Distributions Parameters successfully Enter  ","Screenshot"
	End If
	wait(2)
	OracleFormWindow("Distributions").SelectMenu "File->Save"
	wait(2)
	Verify_Oracle_Status("FRM-40400: Transaction complete: 4 records applied and saved.")
	OracleFormWindow("Distributions").CloseWindow
	wait(2)
	OracleFormWindow("Shipments").CloseWindow
	wait(2)
	
	Purchase_Order_Number = OracleFormWindow("Purchase Orders").OracleTextField("PO, Rev").GetROProperty("value")
	RACK_PutData "Inv_FA_Data", "Purchase_Order_Number", CStr(Purchase_Order_Number)
	Update_Notepad "PO_Number", Purchase_Order_Number
	If  not Purchase_Order_Number = "" Then
	RACK_ReportEvent "Purchase Order", "The Purchase Order '" & Purchase_Order_Number & "' is created sucessfully as expected","Pass"
	else
	RACK_ReportEvent "Purchase Order", "The Purchase Order is not created sucessfully","Fail"
	End if
	OracleFormWindow("Purchase Orders").OracleButton("Approve...").Click
	If  OracleFormWindow("Approve Document").Exist(10) Then
		OracleFormWindow("Approve Document").OracleTabbedRegion("Approval Details").OracleCheckbox("Approval|Submit for Approval").Select
		wait(2)
		OracleFormWindow("Approve Document").OracleTabbedRegion("Approval Details").OracleCheckbox("Approval|Forward").Select
		RACK_ReportEvent "Validation Screenshot", "Approve Document Parameters successfully Enter  ","Screenshot"
		OracleFormWindow("Approve Document").OracleButton("OK").Click
		wait(2)
	End if	
End Function
'#####################################################################################################################
'Function Description   : Function to create a Purchase Order Receipt
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_PO_Receipt()
   	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If OracleListOfValues("Organizations").Exist(90) Then
		OracleListOfValues("Organizations").Select RACK_GetData("Inv_FA_Data", "Organization_Name")
		If OracleFormWindow("Find Expected Receipts").Exist(30) Then
			If not RACK_GetData("Inv_FA_Data", "Purchase_Order_Number") = "" Then
				PO_Number = RACK_GetData("Inv_FA_Data", "Purchase_Order_Number")
			else
				PO_Number = Read_Notepad("PO_Number")
			End If
			OracleFormWindow("Find Expected Receipts").OracleTabbedRegion("Supplier and Internal").OracleTextField("Purchase Order").Enter PO_Number
			OracleFormWindow("Find Expected Receipts").OracleButton("Find").Click
			If OracleFormWindow("Receipt Header").Exist(10)  Then
				OracleFormWindow("Receipt Header").CloseWindow
			End If
			If not RACK_GetData("Inv_FA_Data", "Location") = "" Then
                        OracleFormWindow("Receipts").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Location",RACK_GetData("Inv_FA_Data", "Location")
		    End If
			If not RACK_GetData("Inv_FA_Data", "Subinventory") = "" Then
				        OracleFormWindow("Receipts").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Subinventory",RACK_GetData("Inv_FA_Data", "Subinventory")
						End If
		    If OracleFormWindow("Receipts").Exist(10) Then
               OracleFormWindow("Receipts").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Select Line", true
			   End If
			wait(2)
			If  RACK_GetData("Inv_FA_Data", "Lot_Serial") = "Yes" Then
				OracleFormWindow("Receipts").OracleButton("Lot - Serial").Click
				wait(5)
				OracleFormWindow("Serial Entry").OracleTable("Table").EnterField 1,"Start Serial Number",RACK_GetData("Inv_FA_Data", "Start_Serial")
				wait(2)
                OracleFormWindow("Serial Entry").OracleTable("Table").EnterField 1,"End Serial Number",RACK_GetData("Inv_FA_Data", "End_Serial")
				wait(2)
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
                OracleFormWindow("Serial Entry").OracleButton("Done").Click
			End If
			OracleFormWindow("Receipts").SelectMenu "File->Save"
			wait(2)
			Verify_Oracle_Status("FRM-40400: Transaction complete: 2 records applied and saved.")
		end if
	end if
End Function
'#####################################################################################################################
'Function Description   : Function to Run Cost Manager
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Run_Cost_Manager()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If OracleListOfValues("Organizations").Exist(90) Then
		OracleListOfValues("Organizations").Select RACK_GetData("Inv_FA_Data", "Organization_Name")
    	If OracleFormWindow("Interface Managers").Exist(10) Then
		wait(1)
		'OracleFormWindow("Interface Managers").OracleTextField("Name").SetFocus
		OracleFormWindow("Interface Managers").OracleTable("Table").SetFocus 1,"Name"
		wait(1)
		OracleFormWindow("Interface Managers").SelectMenu "Tools->Launch Manager"
		wait(5)
			If OracleFormWindow("Launch Inventory Managers").Exist(10) Then
				OracleFormWindow("Launch Inventory Managers").OracleButton("Submit_2").Click
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				wait(2)
				Handle_Oracle_Notification_Forms("OK")
				'wait(2)
				'OracleFormWindow("Interface Managers").CloseWindow
				wait(2)
				Notification_Message = Get_Oracle_Notification_Form_Message()
				'msgbox Notification_Message
				Arr=split(Notification_Message, " ")
				Content = Arr(5)
				Asset_Through_Inventory_Request_ID_Arr = Split(Content,")")
				Asset_Through_Inventory_Request_ID = Asset_Through_Inventory_Request_ID_Arr(0)
				'msgbox Asset_Through_Inventory_Request_ID
				Update_Notepad "Asset_Through_Inventory_Request_ID", Asset_Through_Inventory_Request_ID
				Handle_Oracle_Notification_Forms("OK")
				wait(10)
				Verify_Request_Status(Asset_Through_Inventory_Request_ID)
			End if
		End If
	End if
End Function
'#####################################################################################################################
'Function Description   : Function to Create_Accounting_Cost_Manager
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_Accounting_Cost_Manager()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If OracleFlexWindow("Parameters").Exist(120) Then
		OracleFlexWindow("Parameters").OracleTextField("Ledger").Enter RACK_GetData("Inv_FA_Data","Ledger")
		wait(2) 
		OracleFlexWindow("Parameters").OracleTextField("Report").Enter RACK_GetData("Inv_FA_Data","Report")
		wait(2)
		OracleFlexWindow("Parameters").OracleTextField("Include User Transaction").Enter RACK_GetData("Inv_FA_Data","User_Transaction")
		wait(2)
		OracleFlexWindow("Parameters").OracleButton("OK").Click
		wait(2)
	End if
	OracleFormWindow("Create Accounting").OracleButton("Submit").Click
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	Notification_Message = Get_Oracle_Notification_Form_Message()
	Arr=split(Notification_Message, " ")
	Content = Arr(6)
	Asset_Through_Inventory_Request_ID_Arr = Split(Content,")")
	Asset_Through_Inventory_Request_ID = Asset_Through_Inventory_Request_ID_Arr(0)
	Update_Notepad "Asset_Through_Inventory_Request_ID", Asset_Through_Inventory_Request_ID
	Handle_Oracle_Notification_Forms("No")
	Verify_Request_Status(Asset_Through_Inventory_Request_ID)
End function
'#####################################################################################################################
'Function Description   : Function to Run_Create_Accounting_Receiving
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Create_Accounting_Receiving()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	wait(2)	
	OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Inv_FA_Data","Request_Name")
	wait(2)	
	If OracleFlexWindow("Parameters").Exist(10) Then
		OracleFlexWindow("Parameters").OracleTextField("Ledger").Enter RACK_GetData("Inv_FA_Data","Ledger") 
		wait(2)							
		OracleFlexWindow("Parameters").OracleTextField("Report").Enter RACK_GetData("Inv_FA_Data","Report")
		wait(2)	
		OracleFlexWindow("Parameters").OracleTextField("Include User Transaction").Enter RACK_GetData("Inv_FA_Data","User_Transaction")
		wait(2)	
		OracleFlexWindow("Parameters").OracleButton("OK").Click
		wait(2)	
	End if
	OracleFormWindow("Submit Request").OracleButton("Submit").Click
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	Notification_Message = Get_Oracle_Notification_Form_Message()
	Arr=split(Notification_Message, " ")
	Content = Arr(6)
	Asset_Through_Inventory_Request_ID_Arr = Split(Content,")")
	Asset_Through_Inventory_Request_ID = Asset_Through_Inventory_Request_ID_Arr(0)
	Update_Notepad "Asset_Through_Inventory_Request_ID", Asset_Through_Inventory_Request_ID
	Handle_Oracle_Notification_Forms("No")
	Verify_Request_Status(Asset_Through_Inventory_Request_ID)
End function
'#####################################################################################################################
'Function Description   : Function to Interface_Move_transaction_Request
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Interface_Move_transaction_Request()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If OracleFormWindow("Submit a New Request").Exist(120) Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(10) Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Inv_FA_Data", "Request_Name")
			End If
			If OracleFlexWindow("Parameters").Exist(5)Then
                If not RACK_GetData("Inv_FA_Data", "Item_Number") = "" Then
					OracleFlexWindow("Parameters").OracleTextField("Inventory Item").Enter RACK_GetData("Inv_FA_Data", "Item_Number")
				End If
				OracleFlexWindow("Parameters").OracleButton("OK").Click
			End If
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
		End if
	'End if
End Function

'#####################################################################################################################
'Function Description   : Function to IInterface_Fixed_Assets_Request
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Interface_Fixed_Assets_Request()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If OracleFormWindow("Submit a New Request").Exist(120) Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(10) Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Inv_FA_Data", "Request_Name")
			End If
			If OracleFlexWindow("Parameters").Exist(5)Then
                If not RACK_GetData("Inv_FA_Data", "Item_Number") = "" Then
					        OracleFlexWindow("Parameters").OracleTextField("Inventory Item").Enter RACK_GetData("Inv_FA_Data", "Item_Number")
				End If
				OracleFlexWindow("Parameters").OracleButton("OK").Click
			End If
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
		End If
	'End If
	'End If
End Function
'#####################################################################################################################
'Function Description   : Function to Run Post  Mass Addtions
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Post_Mass_Addtions()
    Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_38").Link("Run").Click
	wait(60)
	If OracleFormWindow("Submit a New Request").Exist(10)  Then
		wait(2)
        OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
		wait(2)
        OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	    wait(2)
		'If OracleFormWindow("Submit Request Set").Exist(10)  Then
            OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("Inv_FA_Data","Request_Name")
			wait(5)
			'If  OracleFlexWindow("Parameters").Exist(10) Then
				OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
				OracleFlexWindow("Parameters").OracleTextField("Book").Enter RACK_GetData("Inv_FA_Data", "Asset_Book")
				wait(2)
				OracleFlexWindow("Parameters").Approve
				wait(5)
				OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 2,"Parameters"
				wait(2)
                OracleFlexWindow("Parameters").Approve
				wait(2)
			OracleFormWindow("Submit Request Set").OracleButton("Submit").Click
			RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		Notification_Message = Get_Oracle_Notification_Form_Message()
		Arr=split(Notification_Message, " ")
		Content = Arr(6)
		Payment_Request_ID_Arr = Split(Content,")")
		Payment_Request_ID = Payment_Request_ID_Arr(0)
		Update_Notepad "Payment_Request_ID", Payment_Request_ID
		Handle_Oracle_Notification_Forms("No")
		Verify_Request_Status(Payment_Request_ID)
		else
			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
End Function

'#####################################################################################################################
'Function Description   : Function for Sub Inventory Transfer
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Sub_Inventory_Transfer()
   	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If OracleListOfValues("Organizations").Exist(90) Then
		OracleListOfValues("Organizations").Select RACK_GetData("Inv_FA_Data", "Organization_Name")
		wait(2)
		If OracleFormWindow("Subinventory Transfers").Exist(20) Then
			OracleFormWindow("Subinventory Transfers").OracleTextField("Transaction|Type").Enter RACK_GetData("Inv_FA_Data", "Transaction_Type")
			RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			wait(2)
			OracleFormWindow("Subinventory Transfers").OracleButton("Transaction Lines").Click
			wait(2)
			If OracleFormWindow("Subinventory Transfer").Exist(20) Then
                OracleFormWindow("Subinventory Transfer").OracleTable("Table").EnterField 1,"Item", RACK_GetData("Inv_FA_Data", "Subinventory_item")
				wait(2)
				  OracleFormWindow("Subinventory Transfer").OracleTable("Table").EnterField 1,"Subinventory", RACK_GetData("Inv_FA_Data", "Subinventory")
				 wait(2)
			        OracleFormWindow("Subinventory Transfer").OracleTable("Table").SetFocus 1,"To Subinv"
					OracleFormWindow("Subinventory Transfer").OracleTable("Table").EnterField 1,"To Subinv", RACK_GetData("Inv_FA_Data", "ToSubinventory")
					wait(2)
				    OracleFormWindow("Subinventory Transfer").OracleTable("Table").EnterField 1,"To Locator", RACK_GetData("Inv_FA_Data", "Subinventory_Locator")
					wait(2)
				    OracleFormWindow("Subinventory Transfer").OracleTable("Table_2").SetFocus 1,"Quantity"
				    OracleFormWindow("Subinventory Transfer").OracleTable("Table_3").EnterField 1,"Quantity", RACK_GetData("Inv_FA_Data", "Subinventory_Quantity")
                    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
					wait(2)
				If OracleFormWindow("Subinventory Transfer").OracleButton("Lot / Serial").Exist(20) Then
					OracleFormWindow("Subinventory Transfer").OracleButton("Lot / Serial").Click                            						
					If  OracleFormWindow("Serial Entry").OracleTable("Table").Exist(10) Then
						OracleFormWindow("Serial Entry").OracleTable("Table").OpenDialog 1,"Start Serial Number"
						OracleListOfValues("Serial Numbers").Find "%"
						OracleListOfValues("Serial Numbers").Select 1
						RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"									
						OracleFormWindow("Serial Entry").OracleButton("Done").Click
					end if
					OracleFormWindow("Subinventory Transfer").SelectMenu "File->Save"
					'Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
					wait(2)
					status = OracleStatusLine("OracleStatusLine").GetROProperty("message")
					If  InStr(1,status,"saved") > 0 Then
						RACK_ReportEvent "Subinventory transfer", "The Subinventory transfer is successfull" ,"Pass"
					Else
						RACK_ReportEvent "Subinventory transfer", "The Subinventory transfer is not successfull" ,"Fail"
					End If
				End If
			End If
		End If
	End If
End Function

'#####################################################################################################################
'Function Description   : Function for Subassembly_Build
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Subassembly_Build()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Configure Subassembly")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	wait(5)
	If Parent_desc.WebList("InvOrganizationChoice").Exist(20) then
	    Parent_desc.WebList("InvOrganizationChoice").Select RACK_GetData("Inv_FA_Data", "Organization_Name")
		wait(5)
		Parent_desc.WebEdit("ItemNo").Set RACK_GetData("Inv_FA_Data", "Item_Number")
		wait(3)
		Parent_desc.WebEdit("SerialNumber").Set "18984_Test"
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"	
		wait(3)
		Parent_desc.WebButton("Next").Click
		wait(5)
		If Parent_desc.WebEdit("PartsAdvTableRN:ChassisSerialNumber").Exist(20) then
			Parent_desc.WebEdit("PartsAdvTableRN:ChassisSerialNumber").Set RACK_GetData("Inv_FA_Data", "Chasis_Serial_Number")
			wait(5)
			Parent_desc.WebEdit("PartsAdvTableRN:PartSerialNumber:1").Set RACK_GetData("Inv_FA_Data", "Part_Serial_Number")
			RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			wait(5)
			Parent_desc.WebButton("Next").Click
			RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			If Parent_desc.WebButton("Create").Exist(20) then
				Parent_desc.WebButton("Create").Click
				wait(15)
				Confirmation_Message = Parent_desc.WebTable("Confirmation").GetCellData(1,1)
				Expected_Message =   "Confirmation"
				Display_Message = Browser("Oracle Applications Home Page").Page("Configure Subassembly").WebTable("Subassembly").GetCellData(1,1)
				If  Confirmation_Message = Expected_Message Then
					RACK_ReportEvent "Subassembly Creation", "The Subassembly is created sucessfully with the message  '" & Display_Message & "'" ,"Pass"
				else
					RACK_ReportEvent "Subassembly Creation", "The Subassembly is not created sucessfully","Fail"
				End If
			else
				RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
			End If
		else
			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
		End If
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
End Function
'#####################################################################################################################
'Function Description   : Function to Rackspace Asset Cost Transfer Process
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Asset_Cost_Transfer_File_Upload()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If Parent_desc.WebFile("uploadFile_oafileUpload").Exist(20) Then
        wait(2)
		Parent_desc.WebFile("uploadFile_oafileUpload").Set RACK_GetData("Inv_FA_Data", "Attachement_Path")
		File_Name = Split(RACK_GetData("Inv_FA_Data", "Attachement_Path"),"\")
		Link_Name = File_Name(ubound(File_Name))
		wait(2)
		Parent_desc.WebButton("Submit").Click
		If Parent_desc.Link("name:="& Link_Name).Exist(20) then
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is uploaded as expected","Pass"
		else
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is not uploaded as expected","Fail"
		End if
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End if
End Function

''#####################################################################################################################
''Function Description   : Function for Run_Rackspace_Cost_Transfer_Process
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 

Function Run_Asset_Cost_Transfer()
    Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_38").Link("Run").Click
	wait(60)
	If OracleFormWindow("Submit a New Request").Exist(60)  Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(60)  Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Inv_FA_Data", "Request_Name")
			wait(2)
            If  OracleFlexWindow("Parameters").Exist(10) Then
            OracleFlexWindow("Parameters").OracleTextField("File Name").Enter RACK_GetData("Inv_FA_Data", "Attachement_Path")
			wait(2)
            OracleFlexWindow("Parameters").Approve
			wait(2)
			End If
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
			'RACK_ReportEvent "Validation Screenshot", "FA Mass Retirement  Parameters successfully Enter   ","Screenshot"
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		End If
		Notification_Message = Get_Oracle_Notification_Form_Message()
		Arr=split(Notification_Message, " ")
		'msgbox Arr(6)
		Content = Arr(6)
		Payment_Request_ID_Arr = Split(Content,")")
		Payment_Request_ID = Payment_Request_ID_Arr(0)
		Update_Notepad "Payment_Request_ID", Payment_Request_ID
		Handle_Oracle_Notification_Forms("No")
		Verify_Request_Status(Payment_Request_ID)
		else
			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
End Function
'#####################################################################################################################
'Function Description   : Function to Upload_Direct_Org_Transfer
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Direct_Org_Transfer_File_Upload()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If Parent_desc.WebFile("uploadFile_oafileUpload").Exist(20) Then
        wait(2)
		Parent_desc.WebFile("uploadFile_oafileUpload").Set RACK_GetData("Inv_FA_Data", "Attachement_Path")
		File_Name = Split(RACK_GetData("Inv_FA_Data", "Attachement_Path"),"\")
		Link_Name = File_Name(ubound(File_Name))
		wait(2)
		Parent_desc.WebButton("Submit").Click
		If Parent_desc.Link("name:="& Link_Name).Exist(20) then
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is uploaded as expected","Pass"
		else
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is not uploaded as expected","Fail"
		end if
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	end if
'	end If
'	end If
End Function

''#####################################################################################################################
''Function Description   : Function for Run_Direct_Org_Transfer
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 

Function Run_Direct_Org_Transfer()
    Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_38").Link("Run").Click
	wait(60)
	If OracleFormWindow("Submit a New Request").Exist(60)  Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(60)  Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Inv_FA_Data", "Request_Name")
			wait(2)
            If  OracleFlexWindow("Parameters").Exist(10) Then
            OracleFlexWindow("Parameters").OracleTextField("File Name").Enter RACK_GetData("Inv_FA_Data", "Attachement_Path")
			wait(2)
            OracleFlexWindow("Parameters").Approve
			wait(2)
			End If
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			'RACK_ReportEvent "Validation Screenshot", "FA Mass Retirement  Parameters successfully Enter   ","Screenshot"
		End If
		Notification_Message = Get_Oracle_Notification_Form_Message()
		Arr=split(Notification_Message, " ")
		'msgbox Arr(6)
		Content = Arr(6)
		Payment_Request_ID_Arr = Split(Content,")")
		Payment_Request_ID = Payment_Request_ID_Arr(0)
		Update_Notepad "Payment_Request_ID", Payment_Request_ID
		Handle_Oracle_Notification_Forms("No")
		Verify_Request_Status(Payment_Request_ID)
		else
			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
End Function

'#####################################################################################################################
'Function Description   : Function to Upload_Account_Alias_Issue
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Account_Alias_Issue_File_Upload()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If Parent_desc.WebFile("uploadFile_oafileUpload").Exist(20) Then
        wait(2)
		Parent_desc.WebFile("uploadFile_oafileUpload").Set RACK_GetData("Inv_FA_Data", "Attachement_Path")
		File_Name = Split(RACK_GetData("Inv_FA_Data", "Attachement_Path"),"\")
		Link_Name = File_Name(ubound(File_Name))
		wait(2)
		Parent_desc.WebButton("Submit").Click
		If Parent_desc.Link("name:="& Link_Name).Exist(20) then
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is uploaded as expected","Pass"
		else
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is not uploaded as expected","Fail"
		End if
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End if
End Function
''#####################################################################################################################
''Function Description   : Function for Run_Account_Alias_Issue
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 

Function Run_Account_Alias_Issue()
    Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_38").Link("Run").Click
	wait(60)
	If OracleFormWindow("Submit a New Request").Exist(60)  Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(60)  Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Inv_FA_Data", "Request_Name")
			wait(2)
            If  OracleFlexWindow("Parameters").Exist(10) Then
            OracleFlexWindow("Parameters").OracleTextField("File Name").Enter RACK_GetData("Inv_FA_Data", "Attachement_Path")
			wait(2)
            OracleFlexWindow("Parameters").Approve
			wait(2)
			End If
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			'RACK_ReportEvent "Validation Screenshot", "FA Mass Retirement  Parameters successfully Enter   ","Screenshot"
		End If
		Notification_Message = Get_Oracle_Notification_Form_Message()
		Arr=split(Notification_Message, " ")
		'msgbox Arr(6)
		Content = Arr(6)
		Payment_Request_ID_Arr = Split(Content,")")
		Payment_Request_ID = Payment_Request_ID_Arr(0)
		Update_Notepad "Payment_Request_ID", Payment_Request_ID
		Handle_Oracle_Notification_Forms("No")
		Verify_Request_Status(Payment_Request_ID)
		else
			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
End Function

'#####################################################################################################################
'Function Description   : Function to Upload_Account_Alias_Receipt
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Account_Alias_Receipt_File_Upload()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If Parent_desc.WebFile("uploadFile_oafileUpload").Exist(20) Then
        wait(2)
		Parent_desc.WebFile("uploadFile_oafileUpload").Set RACK_GetData("Inv_FA_Data", "Attachement_Path")
		File_Name = Split(RACK_GetData("Inv_FA_Data", "Attachement_Path"),"\")
		Link_Name = File_Name(ubound(File_Name))
		wait(2)
		Parent_desc.WebButton("Submit").Click
		If Parent_desc.Link("name:="& Link_Name).Exist(20) then
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is uploaded as expected","Pass"
		else
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is not uploaded as expected","Fail"
		End if
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End if
End Function

''#####################################################################################################################
''Function Description   : Function for Run_Account_Alias_Receipt
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 

Function Run_Account_Alias_Receipt()
    Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_38").Link("Run").Click
	wait(60)
	If OracleFormWindow("Submit a New Request").Exist(60)  Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(60)  Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Inv_FA_Data", "Request_Name")
			wait(2)
            If  OracleFlexWindow("Parameters").Exist(10) Then
            OracleFlexWindow("Parameters").OracleTextField("File Name").Enter RACK_GetData("Inv_FA_Data", "Attachement_Path")
			wait(2)
            OracleFlexWindow("Parameters").Approve
			wait(2)
			End If
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			'RACK_ReportEvent "Validation Screenshot", "FA Mass Retirement  Parameters successfully Enter   ","Screenshot"
		End If
		Notification_Message = Get_Oracle_Notification_Form_Message()
		Arr=split(Notification_Message, " ")
		Content = Arr(6)
		Payment_Request_ID_Arr = Split(Content,")")
		Payment_Request_ID = Payment_Request_ID_Arr(0)
		Update_Notepad "Payment_Request_ID", Payment_Request_ID
		Handle_Oracle_Notification_Forms("No")
		Verify_Request_Status(Payment_Request_ID)
		else
			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
End Function

'#####################################################################################################################
'Function Description   : Function to Create_Accounting_Cost_Manager
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_Accounting_Change_Organization()
   	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If OracleListOfValues("Organizations").Exist(90) Then
		OracleListOfValues("Organizations").Select RACK_GetData("Inv_FA_Data", "Organization_Name")
		wait(2)
		OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List_2").Activate "+  Transactions"
		wait(2)
        OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List_2").Select "    Create Accouting - Cost Mgmt"
        OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List_2").Activate "    Create Accouting - Cost Mgmt"
		wait(2)
		OracleFlexWindow("Parameters").OracleTextField("Ledger").Enter RACK_GetData("Inv_FA_Data","Ledger") 
		wait(2)							
		OracleFlexWindow("Parameters").OracleTextField("Report").Enter RACK_GetData("Inv_FA_Data","Report")
		wait(2)
		OracleFlexWindow("Parameters").OracleTextField("Include User Transaction").Enter RACK_GetData("Inv_FA_Data","User_Transaction")
		wait(2)
		OracleFlexWindow("Parameters").OracleButton("OK").Click
		wait(2)
	End if
	OracleFormWindow("Create Accounting").OracleButton("Submit").Click
    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	Notification_Message = Get_Oracle_Notification_Form_Message()
	Arr=split(Notification_Message, " ")
	Content = Arr(6)
	Asset_Through_Inventory_Request_ID_Arr = Split(Content,")")
	Asset_Through_Inventory_Request_ID = Asset_Through_Inventory_Request_ID_Arr(0)
	Update_Notepad "Asset_Through_Inventory_Request_ID", Asset_Through_Inventory_Request_ID
	Handle_Oracle_Notification_Forms("No")
	Verify_Request_Status(Asset_Through_Inventory_Request_ID)
End function

''#####################################################################################################################
''Function Description   : Function for Run Post Mass Retirements
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 

Function Post_Mass_Retirements()
     Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_38").Link("Run").Click
	wait(60)
	If OracleFormWindow("Submit a New Request").Exist(10)  Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(10)  Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Inv_FA_Data", "Request_Name")
			wait(2)
            If  OracleFlexWindow("Parameters").Exist(10) Then
			wait(2)
            OracleFlexWindow("Parameters").OracleTextField("Book").Enter RACK_GetData("Inv_FA_Data", "Asset_Book")
			wait(2)
            OracleFlexWindow("Parameters").Approve
			wait(2)
			End If
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			'RACK_ReportEvent "Validation Screenshot", "Post Mass Retirements Parameters successfully Enter   ","Screenshot"
		End If
		Notification_Message = Get_Oracle_Notification_Form_Message()
		Arr=split(Notification_Message, " ")
		'msgbox Arr(6)
		Content = Arr(6)
		Payment_Request_ID_Arr = Split(Content,")")
		Payment_Request_ID = Payment_Request_ID_Arr(0)
		Update_Notepad "Payment_Request_ID", Payment_Request_ID
		Handle_Oracle_Notification_Forms("No")
		Verify_Request_Status(Payment_Request_ID)
		else
			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
End Function

''#####################################################################################################################
''Function Description   : Function for Run Rackspace FA Retirement Type Update
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 

Function FA_Retirement_Type_Update()
    Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_38").Link("Run").Click
	wait(60)
	If OracleFormWindow("Submit a New Request").Exist(10)  Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(10)  Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Inv_FA_Data", "Request_Name")
			wait(2)
            If  OracleFlexWindow("Parameters").Exist(10) Then
            OracleFlexWindow("Parameters").OracleTextField("Book Type Code").Enter RACK_GetData("Inv_FA_Data", "Asset_Book")
			wait(2)
            OracleFlexWindow("Parameters").Approve
			wait(2)
			End If
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			'RACK_ReportEvent "Validation Screenshot", "FA Retirement Type Update Parameters successfully Enter   ","Screenshot"
		End If
		Notification_Message = Get_Oracle_Notification_Form_Message()
		Arr=split(Notification_Message, " ")
		'msgbox Arr(6)
		Content = Arr(6)
		Payment_Request_ID_Arr = Split(Content,")")
		Payment_Request_ID = Payment_Request_ID_Arr(0)
		Update_Notepad "Payment_Request_ID", Payment_Request_ID
		Handle_Oracle_Notification_Forms("No")
		Verify_Request_Status(Payment_Request_ID)
		else
			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
'	end If
End Function
'#####################################################################################################################
'Function Description   : Function to Mass Retirement  Process
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Mass_Retirement_File_Upload()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If Parent_desc.WebFile("uploadFile_oafileUpload").Exist(20) Then
        wait(2)
		Parent_desc.WebFile("uploadFile_oafileUpload").Set RACK_GetData("Inv_FA_Data", "Attachement_Path")
		File_Name = Split(RACK_GetData("Inv_FA_Data", "Attachement_Path"),"\")
		Link_Name = File_Name(ubound(File_Name))
		wait(2)
		Parent_desc.WebButton("Submit").Click
		If Parent_desc.Link("name:="& Link_Name).Exist(20) then
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is uploaded as expected","Pass"
		else
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is not uploaded as expected","Fail"
		End if
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End if
End Function

''#####################################################################################################################
''Function Description   : Function for Run Rackspace FA Mass Retirement Process
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 

Function Run_Mass_Retirement_Process()
    Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_38").Link("Run").Click
	wait(60)
	If OracleFormWindow("Submit a New Request").Exist(10)  Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(10)  Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Inv_FA_Data", "Request_Name")
			wait(2)
            If  OracleFlexWindow("Parameters").Exist(10) Then
            OracleFlexWindow("Parameters").OracleTextField("File Name").Enter RACK_GetData("Inv_FA_Data", "Attachement_Path")
			wait(2)
            OracleFlexWindow("Parameters").Approve
			wait(2)
			End If
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			'RACK_ReportEvent "Validation Screenshot", "FA Mass Retirement  Parameters successfully Enter   ","Screenshot"
		End If
		Notification_Message = Get_Oracle_Notification_Form_Message()
		Arr=split(Notification_Message, " ")
		'msgbox Arr(6)
		Content = Arr(6)
		Payment_Request_ID_Arr = Split(Content,")")
		Payment_Request_ID = Payment_Request_ID_Arr(0)
		Update_Notepad "Payment_Request_ID", Payment_Request_ID
		Handle_Oracle_Notification_Forms("No")
		Verify_Request_Status(Payment_Request_ID)
		else
			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
End Function

''#####################################################################################################################
''Function Description   : Function for Run Create Accounting - Assets
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 

Function Create_Accounting_Assets()
    Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_38").Link("Run").Click
	wait(60)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	wait(1)
	OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Inv_FA_Data","Request_Name")
    wait(1)
   OracleFlexWindow("Parameters").OracleTextField("Book Type Code").Enter RACK_GetData("Inv_FA_Data","Asset_Book")
   wait(1)
   OracleFlexWindow("Parameters").OracleTextField("Report").Enter RACK_GetData("Inv_FA_Data","Report")
   wait(1)
   OracleFlexWindow("Parameters").Approve
   wait(1)
	OracleFormWindow("Submit Request").OracleButton("Submit").Click
    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	'RACK_ReportEvent "Validation Screenshot", "Create Accounting - Assets Parameters successfully Enter   ","Screenshot"
	Notification_Message = Get_Oracle_Notification_Form_Message()
	Arr=split(Notification_Message, " ")
	Content = Arr(6)
	Payment_Request_ID_Arr = Split(Content,")")
	Payment_Request_ID = Payment_Request_ID_Arr(0)
	Update_Notepad "Payment_Request_ID", Payment_Request_ID
	Handle_Oracle_Notification_Forms("No")
	Verify_Request_Status(Payment_Request_ID)
End Function

''#####################################################################################################################
''Function Description   : Function for Run_Rackspace_Asset_Retirement_Report
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 

Function Rackspace_Asset_Retirement_Report()
    Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_38").Link("Run").Click
	wait(60)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	wait(1)
	OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Inv_FA_Data","Request_Name")
    wait(2)
   OracleFlexWindow("Parameters").OracleTextField("Book").Enter RACK_GetData("Inv_FA_Data","Asset_Book")
  wait(2)
   OracleFlexWindow("Parameters").OracleTextField("Currency").Enter RACK_GetData("Inv_FA_Data","Currency")
   wait(2)
   OracleFlexWindow("Parameters").OracleTextField("From Period").Enter RACK_GetData("Inv_FA_Data","Period")
   wait(2)
   OracleFlexWindow("Parameters").OracleTextField("To Period").Enter RACK_GetData("Inv_FA_Data","Period")
   wait(2)
   OracleFlexWindow("Parameters").Approve
   wait(1)
	OracleFormWindow("Submit Request").OracleButton("Submit").Click
    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	'RACK_ReportEvent "Validation Screenshot", "Create Accounting - Assets Parameters successfully Enter   ","Screenshot"
	Notification_Message = Get_Oracle_Notification_Form_Message()
	Arr=split(Notification_Message, " ")
	Content = Arr(6)
	Payment_Request_ID_Arr = Split(Content,")")
	Payment_Request_ID = Payment_Request_ID_Arr(0)
	Update_Notepad "Payment_Request_ID", Payment_Request_ID
	Handle_Oracle_Notification_Forms("No")
	Verify_Request_Status(Payment_Request_ID)
End Function

''#####################################################################################################################
''Function Description   : Function for Run Rackspace Asset Additions & Adjustments Report
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 

Function Asset_Additions_Adjustments_Report()
    Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home_38").Link("Run").Click
	wait(60)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	wait(1)
	OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Inv_FA_Data","Request_Name")
    wait(2)
   OracleFlexWindow("Parameters").OracleTextField("Book").Enter RACK_GetData("Inv_FA_Data","Asset_Book")
  wait(2)
   OracleFlexWindow("Parameters").OracleTextField("Set of Books Currency").Enter RACK_GetData("Inv_FA_Data","Currency")
   wait(2)
   OracleFlexWindow("Parameters").OracleTextField("Period From").Enter RACK_GetData("Inv_FA_Data","Period")
   wait(2)
   OracleFlexWindow("Parameters").OracleTextField("Period To").Enter RACK_GetData("Inv_FA_Data","Period")
   wait(2)
   RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
   OracleFlexWindow("Parameters").Approve
   wait(1)
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

'#####################################################################################################################
'Function Description   : Function for Run Rackspace On Hand by Serial Number Report
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Hand_By_Serial_Number_Report()
   	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If OracleListOfValues("Organizations").Exist(90) Then
		OracleListOfValues("Organizations").Select RACK_GetData("Inv_FA_Data", "Organization_Name")
		wait(2)
'		OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List_2").Activate "+  Transactions"
'		wait(2)
'        OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List_2").Select "+  Reports"
'		wait(2)
'        OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List_2").Activate "+  Reports"
'		wait(2)
'        OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List_2").Activate "       Transactions"
		If OracleFormWindow("Submit a New Request").Exist(10)  Then
			OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		End if
		If OracleFormWindow("Transaction Reports").Exist(20)  Then
			OracleFormWindow("Transaction Reports").OracleTextField("Run this Request...|Name_3").Enter RACK_GetData("Inv_FA_Data", "Request_Name")
			wait(2)
			If OracleFlexWindow("Parameters").Exist(10)  Then
            OracleFlexWindow("Parameters").Approve
			wait(2)
			End if
			OracleFormWindow("Transaction Reports").OracleButton("Submit_2").Click
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			Notification_Message = Get_Oracle_Notification_Form_Message()
			Arr=split(Notification_Message, " ")
			Content = Arr(6)
			Mass_Sub_Inv_Transfer_Request_ID_Arr = Split(Content,")")
			Mass_Sub_Inv_Transfer_Request_ID = Mass_Sub_Inv_Transfer_Request_ID_Arr(0)
			Update_Notepad "Mass_Sub_Inv_Transfer_Request_ID", Mass_Sub_Inv_Transfer_Request_ID
			Handle_Oracle_Notification_Forms("No")
			Verify_Request_Status(Mass_Sub_Inv_Transfer_Request_ID)
		End if
	End if
End function

'#####################################################################################################################
'Function Description   : Function for Asset_Depreciation
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Run_Depreciation()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	If OracleFormWindow("Run Depreciation").Exist(90) Then
		OracleFormWindow("Run Depreciation").OracleTextField("Book").OpenDialog
		wait(2)
		'OracleFormWindow("Run Depreciation").OracleCheckbox("Close Period").Set "ON"
		OracleFormWindow("Run Depreciation").OracleCheckbox("Close Period").Select
		wait(2)
		OracleFormWindow("Run Depreciation").OracleButton("Run").Click
		Notification_Message = Get_Oracle_Notification_Form_Message()
		Arr=split(Notification_Message, " ")
		Concurrent_Request_ID = Arr(3)
		msgbox Concurrent_Request_ID
		Handle_Oracle_Notification_Forms("OK")
		wait(2)
	End if
	Verify_Request_Status(Concurrent_Request_ID)
End Function

''#####################################################################################################################
''Function Description   : Function for Run_Rackspace_Asset_Retirement_Report
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 

   Function Verify_Assets_Transaction()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	wait(90)
 If not RACK_GetData("Inv_FA_Data","Asset_Number") = "" Then
	 OracleFormWindow("Find Assets").OracleTextField("Asset Number").Enter RACK_GetData("Inv_FA_Data","Asset_Number")
	 End If
	  If not RACK_GetData("Inv_FA_Data","Serial_Number") = "" Then
	 OracleFormWindow("Find Assets").OracleTextField("Serial Number").Enter RACK_GetData("Inv_FA_Data","Serial_Number")
	 End If
	 OracleFormWindow("Find Assets").OracleTextField("Dates in Service From").Enter RACK_GetData("Inv_FA_Data","From_Date")
	 wait(2)
	 OracleFormWindow("Find Assets").OracleTextField("Dates in Service To").Enter RACK_GetData("Inv_FA_Data","From_Date")
	 wait(2)
     RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	 OracleFormWindow("Find Assets").OracleButton("Find").Click
	 Wait(10)
	 RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	 If OracleFormWindow("Assets").Exist(20) Then
	 OracleFormWindow("Assets").OracleButton("Financial Inquiry").Click
	 End If
	 wait(5)
    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	 wait(5)
End Function
'#####################################################################################################################
'Function Description   : Function to update a Item
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Verify_Material_Transaction()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Suppliers")
	Select_Link(RACK_GetData("Inv_FA_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("Inv_FA_Data", "Functionality_Link"))
	wait(10)
	If OracleListOfValues("Organizations").Exist(90) Then
		OracleListOfValues("Organizations").Select RACK_GetData("Inv_FA_Data", "Organization_Name")
		End If
	 If not RACK_GetData("Inv_FA_Data","Item_Number") = "" Then
        OracleFormWindow("Find Material Transactions").OracleTextField("Item").Enter RACK_GetData("Inv_FA_Data","Item_Number")
End If
        'OracleFormWindow("Find Material Transactions").OracleTextField("Transaction Dates").Enter RACK_GetData("Inv_FA_Data","From_Date")
		OracleFormWindow("Find Material Transactions").OracleTextField("Transaction Dates").Enter RACK_GetData("Inv_FA_Data","From_Date")
		wait(2)
        'OracleFormWindow("Find Material Transactions").OracleTextField("To Date: Transaction Dates").Enter RACK_GetData("Inv_FA_Data","From_Date")
		OracleFormWindow("Find Material Transactions").OracleTextField("To Date: Transaction Dates").Enter RACK_GetData("Inv_FA_Data","To_Date")
        'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(5)
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
        OracleFormWindow("Find Material Transactions").OracleButton("Find").Click
		wait(5)
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
        'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(5)
        OracleFormWindow("Material Transactions").OracleButton("Lot / Serial").Click
        'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(10)
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(5)
End Function

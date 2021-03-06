
'#####################################################################################################################
'Function Description   : Function to create a Supplier
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Create_Supplier()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Suppliers")
	Select_Link(RACK_GetData("PO_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("PO_Data", "Functionality_Link"))
	If Parent_desc.WebButton("Create Supplier").Exist(20) then
		Parent_desc.WebButton("Create Supplier").Click
		wait(5)
		OrganizationName = RACK_GetData("PO_Data", "Organization_Name")
		AddressName = RACK_GetData("PO_Data", "AddressName")
		Parent_desc.WebEdit("organization_name").Set OrganizationName
		Parent_desc.WebButton("Apply").Click
		'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	End if
	If Parent_desc.WebButton("Save").Exist(20) then
		Parent_desc.WebButton("Save").Click
		wait(5)
        RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	End if
	If Parent_desc.WebButton("Save").Exist(10) then
		Parent_desc.Link("Address Book").Click
	End if
	If Parent_desc.WebButton("Create").Exist(20) then
        Parent_desc.WebButton("Create").Click
		End If
	If  Parent_desc.WebButton("Continue").Exist(20)Then
		Parent_desc.WebEdit("Address_Line1_Create").Set RACK_GetData("PO_Data", "AddressLine1")
		wait(2)
		Parent_desc.WebEdit("City_Create").Set RACK_GetData("PO_Data", "City")
		wait(2)
		Parent_desc.WebEdit("County_Create").Set RACK_GetData("PO_Data", "County")
		wait(2)
		Parent_desc.WebList("State_Create").Select RACK_GetData("PO_Data", "State")
		wait(2)
		Parent_desc.WebEdit("PostalCode_Create").Set RACK_GetData("PO_Data", "PostalCode")
		wait(2)
		 Browser("Oracle Applications Home Page").Page("Create/Update Address").WebCheckBox("purSite").WebEdit("emailAddr").WebEdit("phNumber").Set RACK_GetData("PO_Data", "Phone_Number")
		wait(2)
		 Browser("Oracle Applications Home Page").Page("Create/Update Address").WebCheckBox("purSite").WebEdit("emailAddr").Set RACK_GetData("PO_Data", "Mail_Address")
		wait(2)
        Browser("Oracle Applications Home Page").Page("Create/Update Address").WebCheckBox("purSite").Set "ON"
		wait(2)
        Browser("Oracle Applications Home Page").Page("Create/Update Address").WebCheckBox("paySite").Set "ON"
		wait(2)
		AddressName = RACK_GetData("PO_Data", "AddressName")
		Parent_desc.WebEdit("AddressName_Create").Set AddressName
		wait(5)
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		Parent_desc.WebButton("Continue").Click
	  ' End If
	   wait(2)
	    Browser("Oracle Applications Home Page").Page("Create Address: Site Creation").WebCheckBox("N10:selected:0").WebCheckBox("N10:selected:3").Set "ON"
'	If Parent_desc.WebCheckBox("N10:selected:0").Exist(20)Then
'		Parent_desc.WebCheckBox("N10:selected:3").Set "ON"
		wait(5)
		'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		Parent_desc.WebButton("Apply").Click
		wait(5)
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(5)
        Browser("Oracle Applications Home Page").Page("Address Book").Link("Invoice Management").Link("Banking Details").Click
        wait(5)
		If Parent_desc.WebButton("Create").Exist(10) then
            Parent_desc.WebButton("Create").Click
			wait(5)
			Browser("Oracle Applications Home Page").Page("Create Bank Account").WebEdit("BranchNameSelect").Set RACK_GetData("PO_Data", "Branch_Name")
			wait(2)
			Browser("Oracle Applications Home Page").Page("Create Bank Account").WebEdit("AcctNumber").Set RACK_GetData("PO_Data", "Account_Number")
            wait(2)
			RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
            Browser("Oracle Applications Home Page").Page("Create Bank Account_3").WebButton("Apply").Click
			wait(5)
            Browser("Oracle Applications Home Page").Page("Bank Accounts").WebButton("Save").Click
			wait(5)
			RACK_ReportEvent "Supplier Creation", "'The Supplier is Created successfully with the message - '" & Confirmation_Message & "'." ,"Pass"
			End If
			End If
End Function

'#####################################################################################################################
'Function Description   : Function to create a Requestion Catalog
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_Catalog_Requisition()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Requisition")
	Select_Link(RACK_GetData("PO_Data", "Responsibility_Link"))
	wait(5)
	'Select_Link(RACK_GetData("PO_Data", "Functionality_Link"))
	'wait(2)
	Browser("name:=.*").Page("title:=.*").Link("name:=Rackspace Catalog").Click
	wait(2)
	'Browser("name:=.*").Page("title:=.*").Link("name:="& RACK_GetData("PO_Data", "Inventory_Group")).Click
	'Select_Link(RACK_GetData("PO_Data", "Inventory_Group"))
	wait(2)
	'Browser("name:=.*").Page("title:=.*").Link("name:="& RACK_GetData("PO_Data", "Inventory_Type")).Click
	'Select_Link(RACK_GetData("PO_Data", "Inventory_Type"))
	Parent_desc.WebEdit("SearchTextInput").Set RACK_GetData("PO_Data", "Inventory_Code")
	Parent_desc.WebButton("Go").Click
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	If Parent_desc.WebButton("Add to Cart").Exist(5) then
		Parent_desc.WebButton("Add to Cart").Click
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(2)
	end if
	If Parent_desc.WebButton("View Cart and Checkout").Exist(5) then
		Parent_desc.WebButton("View Cart and Checkout").Click
		wait(2)
	end if
	If Parent_desc.WebButton("Checkout").Exist(5) then
		Browser("Oracle Applications Home Page").page("Oracle iProcurement: Checkout").WebEdit("Quantity").Set RACK_GetData("PO_Data", "Inventory_Quantity")
		wait(2)
		Parent_desc.WebButton("Checkout").Click
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(2)
	end if
	If  not RACK_GetData("PO_Data", "Requester") = ""Then
		Browser("Oracle Applications Home Page").Page("Oracle iProcurement: Checkout_3").WebEdit("Requester").Set "%%"
		Browser("Oracle Applications Home Page").Page("Oracle iProcurement: Checkout_3").Image("Search for Requester").Click
		Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame").WebEdit("searchText").Set RACK_GetData("PO_Data", "Requester")
		Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame").WebButton("Go").Click
		Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame_").Image("Quick Select").Click
	End If
	If Parent_desc.WebButton("Next").Exist(5) then
		Parent_desc.WebButton("Next").Click
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(2)
	end if
	If Parent_desc.WebButton("Manage Approvals").Exist(5) then
		If  RACK_GetData("PO_Data", "Manage_Approval") <> ""Then
			Approvals = RACK_GetData("PO_Data", "Manage_Approval")
			ApprovalArr = split(Approvals,";")
			Acount = RACK_GetData("PO_Data", "Approval_Count")
			 For  i = 1 to Acount
				Browser("Oracle Applications Home Page").Page("Oracle iProcurement: Checkout_3").WebButton("Manage Approvals").Click
				wait(5)
				Browser("Oracle Applications Home Page").Page("Checkout: Manage Approvals").WebEdit("NewApproverText").Set ApprovalArr(i)
				Browser("Oracle Applications Home Page").Page("Checkout: Manage Approvals").Image("Search for Approver").Click
				wait(5)
				Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame").WebEdit("searchText").Set ApprovalArr(i)
				Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame").WebButton("Go").Click
				Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame").Image("Quick Select").Click
				wait(5)
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				Browser("Oracle Applications Home Page").Page("Checkout: Manage Approvals").WebButton("Submit").Click
				wait(2)
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				wait(30)
				Next
		End If
		'If Parent_desc.WebButton("Next").Exist(5) then
		Parent_desc.WebButton("Next").Click
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(2)
	end if
	If Parent_desc.WebButton("Submit").Exist(5) then
		Parent_desc.WebButton("Submit").Click
		wait(2)
	end if

	If Parent_desc.WebButton("Continue Shopping").Exist(20) then
			Message = Parent_desc.WebElement("ConfirmationRequisition").GetROProperty("innertext")
            Arr=split (Message, " ")
			Req_Num = Arr(1)
			'Msgbox Req_Num
			RACK_PutData "PO_Data", "Request_Number", Arr(1)
			Update_Notepad "Request_Number", Req_Num
			RACK_ReportEvent "Requisition Creation", "The Requistion is created sucessfully with the message '" & Message & "'","Pass"
	else
			RACK_ReportEvent "Requisition Creation", "The Requisition is not created sucessfully","Fail"
'			end If
	End If
End Function


'#####################################################################################################################
'Function Description   : Function to create a Non Requestion Catalog
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_NonCatalog_Requisition()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Requisition")
	Select_Link(RACK_GetData("PO_Data", "Responsibility_Link"))
	wait(5)
	'Select_Link(RACK_GetData("PO_Data", "Functionality_Link"))
'	Browser("name:=.*").Page("title:=.*").Link("name:=Non-Catalog Request").Click
	Browser("name:=.*").Page("title:=.*").Link("name:=   Rackspace Catalog ").Click
	Browser("Oracle Applications Home Page").Page("Oracle iProcurement: Shop_3").Link("RS US NON CATALOG SMART").Click
	wait(2)
    If Parent_desc.WebEdit("ItemDescription").Exist(20) then
		Parent_desc.WebEdit("ItemDescription").set RACK_GetData("PO_Data", "Request_ItemDescription")
		Parent_desc.WebEdit("Category").click
		'WshShell.SendKeys RACK_GetData("PO_Data", "Request_Category")
		'WshShell.SendKeys "{TAB}"
    	wait(10)
		Parent_desc.WebEdit("Quantity").set RACK_GetData("PO_Data", "Request_Quantity")
		wait(5)
		Parent_desc.WebEdit("UnitPrice").set RACK_GetData("PO_Data", "Request_UnitPrice")
		wait(5)
		Parent_desc.WebList("Currency").Select RACK_GetData("PO_Data", "Request_Currency")
		wait(2)
		'RACK_ReportEvent "Validation Screenshot", "Requisition Parameters successfully Enter   ","Screenshot"
		Parent_desc.WebEdit("SupplierContact").click
		wait(3)
		If not RACK_GetData("PO_Data", "Request_Currency") = "USD" Then
			Parent_desc.WebEdit("RateDate").Set RACK_GetData("PO_Data", "Request_ExchangeRateDate")
		End If
    	wait(2)
		Parent_desc.WebEdit("SupplierContact").click
		wait(1)
		RACK_ReportEvent "Validation Screenshot", "Requisition Parameters successfully Enter   ","Screenshot"
		wait(2)
		Parent_desc.WebButton("Add to Cart").Click
	end if
	If Parent_desc.WebButton("View Cart and Checkout").Exist(10) then
		Parent_desc.WebButton("View Cart and Checkout").Click
		wait(2)
	end if
	If Parent_desc.WebButton("Checkout").Exist(10) then
		Parent_desc.WebButton("Checkout").Click
		wait(2)
	end if
	If  not RACK_GetData("PO_Data", "Requester") = ""Then
		Browser("Oracle Applications Home Page").Page("Oracle iProcurement: Checkout_3").WebEdit("Requester").Set "%%"
		Browser("Oracle Applications Home Page").Page("Oracle iProcurement: Checkout_3").Image("Search for Requester").Click
		Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame").WebEdit("searchText").Set RACK_GetData("PO_Data", "Requester")
		Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame").WebButton("Go").Click
		Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame_").Image("Quick Select").Click
	End If
	If Parent_desc.WebButton("Edit Lines").Exist(10) then
		Parent_desc.WebButton("Edit Lines").Click
		wait(2)
'	end if
'	Parent_desc.Link("Accounts").Click
'	Parent_desc.Link("Enter Charge Account").Click
'	If Parent_desc.WebEdit("AccountsDistsAdvTable:ChargeAccount").Exist(5) then
'		Parent_desc.WebEdit("AccountsDistsAdvTable:ChargeAccount").Set RACK_GetData("PO_Data", "Request_ChargeAccount")
'		Parent_desc.WebButton("Apply").Click
		wait(5)
		Parent_desc.WebButton("Apply").Click
	end if
	If Parent_desc.WebButton("Next").Exist(5) then
		Parent_desc.WebButton("Next").Click
		wait(2)
	end if
	If Parent_desc.WebButton("Manage Approvals").Exist(5) then
        If  RACK_GetData("PO_Data", "Manage_Approval") <> ""Then
			Approvals = RACK_GetData("PO_Data", "Manage_Approval")
			ApprovalArr = split(Approvals,";")
			Acount = RACK_GetData("PO_Data", "Approval_Count")
			 For  i = 1 to Acount
				Browser("Oracle Applications Home Page").Page("Oracle iProcurement: Checkout_3").WebButton("Manage Approvals").Click
				wait(5)
				Browser("Oracle Applications Home Page").Page("Checkout: Manage Approvals").WebEdit("NewApproverText").Set ApprovalArr(i)
				Browser("Oracle Applications Home Page").Page("Checkout: Manage Approvals").Image("Search for Approver").Click
				wait(5)
				Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame").WebEdit("searchText").Set ApprovalArr(i)
				Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame").WebButton("Go").Click
				Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame").Image("Quick Select").Click
				wait(5)
				Browser("Oracle Applications Home Page").Page("Checkout: Manage Approvals").WebButton("Submit").Click
				wait(5)
			Next
		End If
		Parent_desc.WebButton("Next").Click
		wait(2)
	end if
	If Parent_desc.WebButton("Submit").Exist(5) then
		Parent_desc.WebButton("Submit").Click
		wait(2)
	end if
	If Parent_desc.WebButton("Continue Shopping").Exist(20) then
			Message = Parent_desc.WebElement("ConfirmationRequisition").GetROProperty("innertext")
            Arr=split (Message, " ")
			Req_Num =  Arr(1)
			RACK_PutData "PO_Data", "Request_Number", Arr(1)
			Update_Notepad "Request_Number", Req_Num
			RACK_ReportEvent "Requisition Creation", "The Requistion is created sucessfully with the message '" & Message & "'","Pass"
	else
			RACK_ReportEvent "Requisition Creation", "The Requisition is not created sucessfully","Fail"
	End If
End Function

'#####################################################################################################################
'Function Description   : Function to create a Purchase Order
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_Purchase_Order()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("PO_Data", "Responsibility_Link"))
	wait(10)
	Select_Link(RACK_GetData("PO_Data", "Functionality_Link"))
	If OracleFormWindow("Find Requisition Lines").Exist(90) Then
    	OracleFormWindow("Find Requisition Lines").OracleButton("Clear").Click
		If not RACK_GetData("PO_Data", "Request_Number") = "" Then
			Request_Number = RACK_GetData("PO_Data", "Request_Number")
		else
			Request_Number = Read_Notepad("Request_Number")
		End If
		OracleFormWindow("Find Requisition Lines").OracleTextField("Requisition").Enter Request_Number
        RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		OracleFormWindow("Find Requisition Lines").OracleButton("Find").Click
		If  OracleFormWindow("AutoCreate Documents").Exist(20) Then
            OracleFormWindow("AutoCreate Documents").OracleTable("Table_2").EnterField 1,"Select Line", true
			'OracleFormWindow("AutoCreate Documents").OracleCheckbox("Select Line").Select
			wait(2)
            'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			OracleFormWindow("AutoCreate Documents").OracleButton("Automatic").Click
			wait(2)
			 RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		end if
	End If
	If OracleFormWindow("New Document").Exist(10) Then
		OracleFormWindow("New Document").OracleButton("Create").Click
		If OracleNotification("Caution").Exist(5) Then
			OracleNotification("Caution").OracleButton("OK").Click
		End If
		If  OracleFormWindow("AutoCreate to Purchase").Exist(10)Then
			Purchase_Order_Number = OracleFormWindow("AutoCreate to Purchase").OracleTextField("PO, Rev").GetROProperty("value")
			Update_Notepad "PO_Number", Purchase_Order_Number
			RACK_PutData "PO_Data", "Purchase_Order_Number", CStr(Purchase_Order_Number)
			If  not Purchase_Order_Number = "" Then
				RACK_ReportEvent "Purchase Order", "The Purchase Order '" & Purchase_Order_Number & "' is created sucessfully for/with Request Number '" & Request_Number & "' as expected","Pass"
				If not RACK_GetData("PO_Data", "Supplier_Name") ="" Then
					OracleFormWindow("AutoCreate to Purchase").OracleTextField("Supplier").Enter RACK_GetData("PO_Data", "Supplier_Name")
				End If
            	wait(2)
				If not RACK_GetData("PO_Data", "PO_Quantity") ="" Then
                    OracleFormWindow("AutoCreate to Purchase").OracleTabbedRegion("Lines").OracleTable("Table_2").EnterField 1,"Quantity", RACK_GetData("PO_Data", "PO_Quantity")
					wait(5)
					OracleFormWindow("AutoCreate to Purchase").OracleTabbedRegion("Lines").OracleTable("Table_2").EnterField 1,"Price", RACK_GetData("PO_Data", "PO_Price")
				'	OracleFormWindow("AutoCreate to Purchase").OracleTabbedRegion("Lines").OracleTextField("Quantity").Enter RACK_GetData("PO_Data", "PO_Quantity")
					wait(5)
				End If
                'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				OracleFormWindow("AutoCreate to Purchase").OracleButton("Approve...").Click
				wait(2)
				RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				wait(5)
				OracleFormWindow("Approve Document").OracleButton("OK").Click
'			else
'				RACK_ReportEvent"Request Status", "The Purchase Order is not created sucessfully","Fail"
'			end if
'			If  OracleFormWindow("Approve Document").Exist(10) Then
'				OracleFormWindow("Approve Document").OracleButton("OK").Click
'                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			wait(2)
			End If
		End If
	End If
End Function

'#####################################################################################################################
'Function Description   : Function to create a Blanket Purchase Order
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_Blanket_Purchase_Order()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("PO_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("PO_Data", "Functionality_Link"))
	wait(60)
	If OracleFormWindow("Purchase Orders").OracleTextField("Type").Exist(90) Then
		OracleFormWindow("Purchase Orders").OracleTextField("Type").Enter RACK_GetData("PO_Data", "PO_Type")
		OracleFormWindow("Purchase Orders").OracleTextField("Supplier").SetFocus
		wait(3)
		OracleFormWindow("Purchase Orders").OracleTextField("Supplier").Enter RACK_GetData("PO_Data", "Supplier_Name")
		OracleFormWindow("Purchase Orders").OracleTextField("Site").SetFocus
		wait(2)
        OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table_2").SetFocus 1,"Item"
		'OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Item").SetFocus
		wait(4)
		OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table_2").EnterField 1,"Item",RACK_GetData("PO_Data", "PO_Item")
		'OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Item").Enter RACK_GetData("PO_Data", "PO_Item")
		'OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table_2").SetFocus 1,"Quantity"
		'OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Quantity").SetFocus
		'wait(2)
		'OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table_2").EnterField 1,"Quantity",RACK_GetData("PO_Data", "PO_Quantity")
		'OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Quantity").Enter RACK_GetData("PO_Data", "PO_Quantity")
		wait(2)
		OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table_2").EnterField 1,"Price",RACK_GetData("PO_Data", "PO_Price")
		'OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Price").Enter RACK_GetData("PO_Data", "PO_Price")
		'OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table_2").SetFocus 1,"Need-By"
		'OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Need-By").SetFocus
		'wait(2)
			'OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table_2").EnterField 1,"Need-By",RACK_GetData("PO_Data", "PO_Needby")
			wait(2)
		'OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Need-By").Enter RACK_GetData("PO_Data", "PO_Needby")
		OracleFormWindow("Purchase Orders").SelectMenu "File->Save"
		wait(5)
		Verify_Oracle_Status("FRM-40400: Transaction complete: 2 records applied and saved.")
		Purchase_Order_Number = OracleFormWindow("Purchase Orders").OracleTextField("PO, Rev").GetROProperty("value")
		RACK_PutData "PO_Data", "Purchase_Order_Number", CStr(Purchase_Order_Number)
		Update_Notepad "PO_Number", Purchase_Order_Number
		OracleFormWindow("Purchase Orders").OracleButton("Approve...").Click
		If  OracleFormWindow("Approve Document").Exist(10) Then
			OracleFormWindow("Approve Document").OracleTabbedRegion("Approval Details").OracleCheckbox("Approval|Submit for Approval").Select
			OracleFormWindow("Approve Document").OracleButton("OK").Click
			wait(5)
		end if
		If  not Purchase_Order_Number = "" Then
				RACK_ReportEvent "Purchase Order", "The Purchase Order '" & Purchase_Order_Number & "' is created sucessfully as expected","Pass"
		else
				RACK_ReportEvent "Purchase Order", "The Purchase Order is not created sucessfully","Fail"
		end if
	End If
End Function


'#####################################################################################################################
'Function Description   : Function to create a Purchase Order Receipt
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_PO_Receipt()
   	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
   	'set WshShell = CreateObject("WScript.Shell")
	Select_Link(RACK_GetData("PO_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("PO_Data", "Functionality_Link"))
	wait(10)
	If OracleListOfValues("Organizations").Exist(90) Then
		OracleListOfValues("Organizations").Select RACK_GetData("PO_Data", "Organization_Name")
		If OracleFormWindow("Find Expected Receipts").Exist(30) Then
			If not RACK_GetData("PO_Data", "Purchase_Order_Number") = "" Then
				PO_Number = RACK_GetData("PO_Data", "Purchase_Order_Number")
			else
				PO_Number = Read_Notepad("PO_Number")
			End If
			OracleFormWindow("Find Expected Receipts").OracleTabbedRegion("Supplier and Internal").OracleTextField("Purchase Order").Enter PO_Number
			OracleFormWindow("Find Expected Receipts").OracleButton("Find").Click
			If OracleFormWindow("Receipt Header").Exist(10)  Then
				OracleFormWindow("Receipt Header").CloseWindow
				wait(3)
			End If
		End If
		If OracleFormWindow("Receipts").Exist(10) Then
            OracleFormWindow("Receipts").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Select Line", true
            wait(2)
            OracleFormWindow("Receipts").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Location",RACK_GetData("PO_Data", "Location")
			 'OracleFormWindow("Receipts").OracleTabbedRegion("Lines").OracleTable("Table_2").EnterField
			'OracleFormWindow("Receipts").OracleTabbedRegion("Lines").OracleCheckBox("Select Line").Select
			wait(2)
			If  RACK_GetData("PO_Data", "Lot_Serial") = "Yes" Then
				OracleFormWindow("Receipts").OracleButton("Lot - Serial").Click
'				If OracleFormWindow("Serial Entry").Exist(10) Then
'					'msgbox RACK_GetData("PO_Data", "Start_Serial_Number")
'					OracleFormWindow("Serial Entry").OracleTextField("Start Serial Number").Enter RandomNumber (99,99999)
'					OracleFormWindow("Serial Entry").OracleTextField("End Serial Number").Click
'					wait(5)
'					OracleFormWindow("Serial Entry").OracleButton("Done").Click
'				End If
				wait(5)
				OracleFormWindow("Serial Entry").OracleTable("Table").EnterField 1,"Start Serial Number",RACK_GetData("PO_Data", "Start_Serial")
				wait(2)
                OracleFormWindow("Serial Entry").OracleTable("Table").EnterField 1,"End Serial Number",RACK_GetData("PO_Data", "End_Serial")
				wait(2)
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
                OracleFormWindow("Serial Entry").OracleButton("Done").Click

'                		'OracleFormWindow("Serial Entry").OracleTable("Table").SetFocus 1,"Start Serial Number"
'						OracleFormWindow("Serial Entry").OracleTable("Table").EnterField 1,"Start Serial Number","Test_123"
'						'OracleFormWindow("Serial Entry").OracleTable("Table").SetFocus 1,"End Serial Number"
'						'OracleFormWindow("Serial Entry").OracleTable("Table").SetFocus 2,"Start Serial Number"
'						'OracleFormWindow("Serial Entry").OracleTable("Table").EnterField 2,"Start Serial Number"
'						OracleFormWindow("Serial Entry").OracleTable("Table").EnterField 1,"End Serial Number","Test_132"
'						OracleFormWindow("Serial Entry").OracleButton("Done").Click
			End If
			OracleFormWindow("Receipts").SelectMenu "File->Save"
			wait(2)
			Verify_Oracle_Status("FRM-40400: Transaction complete: 2 records applied and saved.")
		end if
	end if
End Function
		
'#####################################################################################################################
'Function Description   : Function to create a Blanket Release
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Blanket_Release()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("PO_Data", "Responsibility_Link"))
	wait(2)
	Select_Link(RACK_GetData("PO_Data", "Functionality_Link"))
	wait(60)
	If not RACK_GetData("PO_Data", "Purchase_Order_Number") = "" Then
		PO_Number = RACK_GetData("PO_Data", "Purchase_Order_Number")
	else
		PO_Number = Read_Notepad("PO_Number")
	End If
	'OracleFormWindow("Releases").OracleTextField("PO, Rev").Enter PO_Number
	'OracleFormWindow("Releases").OracleTextField("Release").SetFocus
	'OracleFormWindow("Releases").OracleTextField("Release").Enter RACK_GetData("PO_Data", "PO_Release")
	'OracleFormWindow("Releases").OracleTextField("Release").Enter "1"
	wait(2)
	'OracleFormWindow("Releases").OracleTabbedRegion("Shipments").OracleTable("Table").EnterField 1,"Line",OracleFormWindow("Releases").OracleTextField("Po_Line_Num")
	OracleFormWindow("Releases").OracleTabbedRegion("Shipments").OracleTable("Table").EnterField 1,"Line","1"
	wait(2)
	'OracleFormWindow("Releases").OracleTabbedRegion("Shipments").OracleTable("Table").EnterField 1,"Quantity",OracleFormWindow("Releases").OracleTextField("PO_Qty")
	OracleFormWindow("Releases").OracleTabbedRegion("Shipments").OracleTable("Table").EnterField 1,"Quantity","10"
	wait(2)
	'OracleFormWindow("Releases").OracleTabbedRegion("Shipments").OracleTable("Table_2").EnterField 1,"Need-By",OracleFormWindow("Releases").OracleTextField("PO_Need_By_Date")
	OracleFormWindow("Releases").OracleTabbedRegion("Shipments").OracleTable("Table_2").EnterField 1,"Need-By",RACK_GetData("PO_Data", "PO_Need_By_Date")
	'OracleFormWindow("Releases").OracleTabbedRegion("Shipments").OracleTable("Table").OpenDialog 1,"Need-By"
	wait(2)
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	'OracleFormWindow("Releases").OracleButton("Distributions").Click
    'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	OracleFormWindow("Distributions").SelectMenu "File->Save"
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	OracleFormWindow("Distributions").CloseWindow
	OracleFormWindow("Releases").OracleButton("Approve...").Click
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	OracleFormWindow("Approve Document").OracleTabbedRegion("Approval Details").OracleCheckbox("Transmission Methods|E-Mail").Clear
	OracleFormWindow("Approve Document").OracleButton("OK").Click
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	RACK_ReportEvent "Validation Screenshot", "Blanket_Release is created sucessfully ","Screenshot"

End Function
'#####################################################################################################################
'Function Description   : Function to Transfer to GL
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Transfer_to_GL()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("PO_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("PO_Data", "Functionality_Link"))
	If OracleFormWindow("Submit a New Request").Exist(90) Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(10) Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("PO_Data", "Request_Name")
			wait(2)
			If  OracleFlexWindow("Parameters").Exist(10) Then
				OracleFlexWindow("Parameters").OracleTextField("Ledger").Enter RACK_GetData("PO_Data", "Ledger")
				'End_Date = Format_Date("dd-mmm-yyyy")
				OracleFlexWindow("Parameters").OracleTextField("End Date").Enter RACK_GetData("PO_Data", "End_Date")
				OracleFlexWindow("Parameters").OracleButton("OK").Click
				wait(2)
			End If
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		End If
		Notification_Message = Get_Oracle_Notification_Form_Message()
		Arr=split(Notification_Message, " ")
		'msgbox Arr(6)
		Content = Arr(6)
		Payment_Request_ID_Arr = Split(Content,")")
		Payment_Request_ID = Payment_Request_ID_Arr(0)
		'msgbox Payment_Request_ID
		'msgbox Payment_Request_ID
		Update_Notepad "Payment_Request_ID", Payment_Request_ID
		Handle_Oracle_Notification_Forms("No")
	end if
	Verify_Request_Status(Payment_Request_ID)
'	OracleFormWindow("Navigator").SelectMenu "View->Requests"
'	OracleFormWindow("Find Requests").OracleRadioGroup("My Completed Requests").Select "Specific Requests"
'	OracleFormWindow("Find Requests").OracleTextField("Request ID").Enter Payment_Request_ID
'	wait(30)
'	OracleFormWindow("Find Requests").OracleButton("Find").Click
'	If OracleFormWindow("Requests").Exist(20) Then
'		wait(30)
'		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
'		If (OracleFormWindow("Requests").OracleTextField("Phase").GetROProperty("value") = "Completed") Then
'			RACK_ReportEvent "Request Phase", "The Request Phase is correctly displayed as '" & OracleFormWindow("Requests").OracleTextField("Phase").GetROProperty("value")& "'.","Pass"
'		Else
'			RACK_ReportEvent "Request Phase", "The Request Phase is not correctly displayed as Expected. The Expected Value is 'Completed' and displayed Value is '" & OracleFormWindow("Requests").OracleTextField("Phase").GetROProperty("value")& "'.","Fail"
'		End If
'		If (OracleFormWindow("Requests").OracleTextField("Status").GetROProperty("value") = "Normal") Then
'			RACK_ReportEvent "Request Status", "The Request Status is correctly displayed as '" & OracleFormWindow("Requests").OracleTextField("Status").GetROProperty("value")& "'.","Pass"
'		Else
'			RACK_ReportEvent "Request Status", "The Request Status is not correctly displayed as Expected. The Expected Value is 'Normal' and displayed Value is '" & OracleFormWindow("Requests").OracleTextField("Status").GetROProperty("value")& "'.","Fail"
'		End If
'   	End if
End Function

'#####################################################################################################################
'Function Description   : Function to Vacation Rule Delegation
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Vacation_Rule_Delegation()
    set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Workflow")
	Select_Link(RACK_GetData("PO_Data", "Responsibility_Link"))
	wait(2)
	Select_Link(RACK_GetData("PO_Data", "Functionality_Link"))
	If Parent_desc.WebButton("Reassign").Exist(10)Then
		Parent_desc.Link("Vacation Rules").Click
		wait(4)
		Parent_desc.WebButton("Create Rule").Click
		Browser("Vacation Rules").Page("Vacation Rule: Item Type").WebList("NRREITItemTypes").Select "Rackspace PO Approval"
		Parent_desc.WebButton("Next").Click
		wait(2)
		Parent_desc.WebButton("Next").Click
		msgbox Format_Date("dd+1-mmm-yyyy")
		Parent_desc.WebEdit("NRREREndDate").Set Format_Date("dd+1-mmm-yyyy")
		'Parent_desc.WebEdit("NRREREndDate").Set Format_Date("dd-mmm-yyyy")
		Parent_desc.WebEdit("NRRERComments").Set "Cpu Patch Test"
		Parent_desc.WebEdit("wfUserName1").click
		WshShell.SendKeys RACK_GetData("PO_Data", "Assign_To")
		'WshShell.SendKeys "{TAB}"
		'Parent_desc.WebEdit("wfUserName1").Set RACK_GetData("PO_Data", "Assign_To")
		wait(10)
		Parent_desc.WebButton("Apply").Click
        RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(10)
		RACK_ReportEvent "Validation Screenshot", "Rule_Delegation is created sucessfully ","Pass"

	End If
End Function

'#####################################################################################################################
'Function Description   : Function to Freeze the PO
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Close_For_Receiving()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("PO_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("PO_Data", "Functionality_Link"))
	If OracleFormWindow("Find Purchase Orders").OracleTextField("Number").Exist(90) Then
		If not RACK_GetData("PO_Data", "Purchase_Order_Number") = "" Then
			PO_Number = RACK_GetData("PO_Data", "Purchase_Order_Number")
		else
			PO_Number = Read_Notepad("PO_Number")
		End If
		'msgbox PO_Number
		OracleFormWindow("Find Purchase Orders").OracleTextField("Number").Enter PO_Number
		wait(2)
		OracleFormWindow("Find Purchase Orders").OracleButton("Find").Click
		If OracleFormWindow("Purchase Order Headers").Exist(20) Then
			OracleFormWindow("Purchase Order Headers").SelectMenu "Tools->Control"
			wait(2)
			If  OracleFormWindow("Control Document").Exist(90) Then
				OracleFormWindow("Control Document").OracleList("Actions").Select RACK_GetData("PO_Data", "Actions")
				wait(2)
				OracleFormWindow("Control Document").OracleTextField("Reason").Enter RACK_GetData("PO_Data", "Reason")
				OracleFormWindow("Control Document").OracleButton("OK").Click
				wait(2)
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				OracleNotification("Note").Approve
				wait(2)
			End If
	End If
End If
End Function

'#####################################################################################################################
'Function Description   : Function to Close a PO
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Close_PO()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("PO_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("PO_Data", "Functionality_Link"))
	If OracleFormWindow("Find Purchase Orders").OracleTextField("Number").Exist(90) Then
		If not RACK_GetData("PO_Data", "Purchase_Order_Number") = "" Then
			PO_Number = RACK_GetData("PO_Data", "Purchase_Order_Number")
		else
			PO_Number = Read_Notepad("PO_Number")
		End If
		'msgbox PO_Number
		OracleFormWindow("Find Purchase Orders").OracleTextField("Number").Enter PO_Number
		wait(1)
		OracleFormWindow("Find Purchase Orders").OracleButton("Find").Click
		If OracleFormWindow("Purchase Order Headers").Exist(20) Then
			OracleFormWindow("Purchase Order Headers").SelectMenu "Tools->Control"
			'OracleFormWindow("Purchase Order Headers").SelectMenu "Tools->Control"
			wait(2)
			If  OracleFormWindow("Control Document").Exist(90) Then
				OracleFormWindow("Control Document").OracleList("Actions").Select RACK_GetData("PO_Data", "Actions")
				wait(2)
				OracleFormWindow("Control Document").OracleTextField("Reason").Enter RACK_GetData("PO_Data", "Reason")
				OracleFormWindow("Control Document").OracleButton("OK").Click
				wait(2)
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				OracleNotification("Note").Approve
				wait(2)
			End If
	End If
End If
End Function
			
'#####################################################################################################################
'Function Description   : Function to Create Purchase Order
'Input Parameters 	: None
'Return Value    	: None
'#####################################################################################################################

Function Create_Manual_Purchase_Order()
	Select_Link(RACK_GetData("PO_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("PO_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Purchase Orders").OracleTextField("Type").SetFocus
	If not  RACK_GetData("PO_Data", "PO_Type") = "" Then
		OracleFormWindow("Purchase Orders").OracleTextField("Type").Enter RACK_GetData("PO_Data", "PO_Type")
	End If
	OracleFormWindow("Purchase Orders").OracleTextField("Supplier").Enter RACK_GetData("PO_Data", "Supplier_Name")
	If not RACK_GetData("PO_Data", "Supplier_To_Site") = "" Then
		OracleFormWindow("Purchase Orders").OracleTextField("Site").Enter RACK_GetData("PO_Data", "Supplier_To_Site")
	End If
	If not RACK_GetData("PO_Data", "PO_Item") = "" Then
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").SetFocus 1,"Item"
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Item",RACK_GetData("PO_Data", "PO_Item")
	End If
	If not RACK_GetData("PO_Data", "PO_Category") = "" Then
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").SetFocus 1,"Category"
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Category",RACK_GetData("PO_Data", "PO_Category")
			wait(2)
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Description",RACK_GetData("PO_Data", "PO_Description")
	End If
	OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Quantity",RACK_GetData("PO_Data", "PO_Quantity")
	wait(2)
	OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table_3").EnterField 1,"Price",RACK_GetData("PO_Data", "PO_Price")
	wait(2)
	OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table_4").EnterField 1,"Need-By",RACK_GetData("PO_Data", "PO_Need_By_Date")
	wait(2)
	RACK_ReportEvent "Validation Screenshot", "Purchase Orders Parameters successfully Enter  ","Screenshot"
	OracleFormWindow("Purchase Orders").OracleButton("Shipments").Click
	wait(2)
	RACK_ReportEvent "Validation Screenshot", "Shipments Parameters successfully Enter  ","Screenshot"
	OracleFormWindow("Shipments").OracleButton("Distributions").Click
	wait(2)
	If not RACK_GetData("PO_Data", "Po_Charge_Account") = "" Then
		OracleFormWindow("Distributions").OracleTabbedRegion("Destination").OracleTable("Table").SetFocus 1,"PO Charge Account"
		'OracleFlexWindow("Charge Account").Cancel
		OracleFormWindow("Distributions").OracleTabbedRegion("Destination").OracleTable("Table_2").EnterField 1,"PO Charge Account",RACK_GetData("PO_Data", "Po_Charge_Account")
		'RACK_ReportEvent "Validation Screenshot", "Distributions Parameters successfully Enter  ","Screenshot"
	End If
	wait(2)
	OracleFormWindow("Distributions").SelectMenu "File->Save"
	Wait(2)
	Verify_Oracle_Status("FRM-40400: Transaction complete: 4 records applied and saved.")
	OracleFormWindow("Distributions").CloseWindow
	Wait(2)
	OracleFormWindow("Shipments").CloseWindow
	wait(2)
	
	Purchase_Order_Number = OracleFormWindow("Purchase Orders").OracleTextField("PO, Rev").GetROProperty("value")
	RACK_PutData "PO_Data", "Purchase_Order_Number", CStr(Purchase_Order_Number)
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
'Function Description   : Function to Create Purchase Order
'Input Parameters 	: None
'Return Value    	: None
'#####################################################################################################################

Function Reassign_PO_For_Approval()
	If RACK_GetData("PO_Data", "ApprovalData") = "Request" Then
		Approval_Number = Read_Notepad("Request_Number")
	elseif RACK_GetData("PO_Data", "ApprovalData") = "PO" Then
	    wait(10)
		Approval_Number = Read_Notepad("PO_Number")	
	elseif RACK_GetData("PO_Data", "ApprovalData") = "INVOICE" Then
	    wait(10)
		Approval_Number = Read_Notepad("Invoice_Number")	
	else
		Approval_Number = RACK_GetData("PO_Data", "ApprovalData")
	End If
	'msgbox Approval_Number
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Workflow")
	Select_Link(RACK_GetData("PO_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("PO_Data", "Functionality_Link"))

	If Parent_desc.WebButton("Reassign").Exist(10)Then
		'Parent_desc.WebCheckBox("N25:selected:0").Set "ON"
		Browser("name:=.*").Page("title:=.*").Link("name:=.*"& Approval_Number & ".*").Click
		wait(1)
	End If
	If Parent_desc.WebButton("Approve And Forward").Exist(10)Then
		Parent_desc.WebEdit("wfUserName1").click
		WshShell.SendKeys RACK_GetData("PO_Data", "Forward_To")
		WshShell.SendKeys "{TAB}"
		'Parent_desc.WebEdit("wfUserName1").Set RACK_GetData("PO_Data", "Forward_To")
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
'Function Description   : Function to Create Accounting
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_Accounting()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("VERTEX_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("VERTEX_Data", "Functionality_Link"))
    If OracleFormWindow("Submit a New Request").Exist(90) Then
		'OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(10) Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("VERTEX_Data", "Request_Name")
			wait(2)
             'OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
              OracleFlexWindow("Parameters").OracleTextField("Ledger").Enter RACK_GetData("VERTEX_Data", "Ledger")
              OracleFlexWindow("Parameters").OracleTextField("End Date").Enter RACK_GetData("VERTEX_Data", "End_Date")
              OracleFlexWindow("Parameters").OracleTextField("Report").Enter RACK_GetData("VERTEX_Data", "Report")
              OracleFlexWindow("Parameters").Approve
			'End If
			wait(2)
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
			RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			'Handle_Oracle_Notification_Forms("OK")
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

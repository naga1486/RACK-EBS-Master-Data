'##################################################################################################################### 
'Function Description   : Function to Vertex create a Invoice Batches
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Vertex_Create_AP_Invoice()
   	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Vertex_Data", "Responsibility_Link"))
	wait(5)
	Parent_desc.Link("Invoice Batches").click

	If OracleFormWindow("Invoice Batches").OracleTable("Table").Exist(120) Then
        OracleFormWindow("Invoice Batches").OracleTable("Table").EnterField 1,"Batch Name", RACK_GetData("Vertex_Data", "BatchName")
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(2)
		OracleFormWindow("Invoice Batches").OracleButton("Invoices").Click
		End If
            If OracleFormWindow("Invoice Workbench").OracleTable("Table").Exist(10) Then
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Type", RACK_GetData("Vertex_Data", "Invoice_Type")
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Trading Partner", RACK_GetData("Vertex_Data", "Supplier_Name")
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Date", RACK_GetData("Vertex_Data", "Invoice_Date") 
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Num", RACK_GetData("Vertex_Data", "Invoice_Number")
			Update_Notepad "Invoice_Number", Invoice_Number
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Amount", RACK_GetData("Vertex_Data", "Invoice_Amount")
			wait(2)
			If not  RACK_GetData("Vertex_Data", "Requester") = "" Then
			OracleFormWindow("Invoice Workbench").SelectMenu "Folder->Folder Tools"
			wait(2)
            OracleFormWindow("Folder Tools").OracleButton("Show Field...").Click
			wait(2)
            OracleListOfValues("Show Field").Select "Requester                                                                                                                                                                                                                  *REQUESTER_NAME"
            OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 1,"Requester", RACK_GetData("Vertex_Data", "Requester")
			wait(2)
			OracleFormWindow("Folder Tools").CloseWindow
            'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		   End If
	       wait(2)
  			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 1,"Amount",RACK_GetData("Vertex_Data", "Item_Amount")
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 2,"Type",RACK_GetData("Vertex_Data", "Type")
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 2,"Amount",RACK_GetData("Vertex_Data", "Tax_Amount")
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 1,"Type","Item"
			 wait(2)
             RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleButton("Distributions").Click
			 wait(2)
			OracleFormWindow("Distributions").OracleTable("Table").EnterField 1,"Amount",RACK_GetData("Vertex_Data", "Item_Amount")
			 wait(2)
           OracleFormWindow("Distributions").OracleTable("Table").EnterField 1,"Account",RACK_GetData("Vertex_Data","Account_Number") 
		    wait(2)
		   OracleFormWindow("Distributions").SelectMenu "File->Save"
		   Verify_Oracle_Status("FRM-40400:Transaction complete: 2 records applied and saved.")
		   'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		   OracleFormWindow("Distributions").CloseWindow
		   RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		    wait(5)
				 End If
				 If RACK_GetData("Vertex_Data","Calculate")= "Yes" Then
					 OracleFormWindow("Invoice Workbench").SelectMenu "Tools->Calculate Tax"
					 wait(2)
					 RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
					 End If
				'If OracleFormWindow("Invoice Workbench").Exist(10)  Then
				If RACK_GetData("Vertex_Data","Actions")= "Yes" Then
				OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
				wait(3)
				OracleFormWindow("Invoice Actions").OracleCheckbox("Validate").Select
				RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				wait(2)
				OracleFormWindow("Invoice Actions").OracleButton("OK").Click
				wait(2)
				OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
				wait(2)
				If RACK_GetData("Vertex_Data","Create_Accounting")= "Yes" Then
                OracleFormWindow("Invoice Actions").OracleCheckbox("Create Accounting").Select
				OracleFormWindow("Invoice Actions").OracleRadioGroup("Draft").Select "Final"
				OracleFormWindow("Invoice Actions").OracleCheckbox("Initiate Approval").Select
				End If
                If RACK_GetData("Vertex_Data","Force_Approval")= "Yes" Then
					OracleFormWindow("Invoice Actions").OracleCheckbox("Force Approval").Select
					End If
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
                OracleFormWindow("Invoice Actions").OracleButton("OK").Click
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
OracleNotification("Note").Approve
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
wait(5)
End If
End Function
'#####################################################################################################################
'Function Description   : Function to Vertex Find AP Invoice
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Vertex_Find_AP_Invoice()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Vertex_Data", "Responsibility_Link"))
	wait(5)
	Parent_desc.Link("Invoice Batches").click
	wait(60)
	    OracleFormWindow("Invoice Batches").SelectMenu "View->Find..."
		wait(2)
		OracleFormWindow("Find Invoice Batches").OracleTextField("Names").Enter RACK_GetData("Vertex_Data", "BatchName")
		wait(2)
		OracleFormWindow("Find Invoice Batches").OracleButton("Find").Click
		wait(2)
		OracleFormWindow("Invoice Batches").OracleButton("Invoices").Click
		wait(2)
		OracleFormWindow("Invoice Workbench").OracleTable("Table_2").SetFocus 1,"Type"
		OracleFormWindow("Invoice Workbench").SelectMenu "File->New"
        If OracleFormWindow("Invoice Workbench").OracleTable("Table_2").Exist(10) Then
			OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Type", RACK_GetData("Vertex_Data", "Invoice_Type")
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Trading Partner", RACK_GetData("Vertex_Data", "Supplier_Name")
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Invoice Date", RACK_GetData("Vertex_Data", "Invoice_Date") 
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Invoice Num", RACK_GetData("Vertex_Data", "Invoice_Number")
			Update_Notepad "Invoice_Number", Invoice_Number
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Invoice Amount", RACK_GetData("Vertex_Data", "Invoice_Amount")
			wait(2)
			If not  RACK_GetData("Vertex_Data", "Requester") = "" Then
			OracleFormWindow("Invoice Workbench").SelectMenu "Folder->Folder Tools"
			wait(2)
            OracleFormWindow("Folder Tools").OracleButton("Show Field..._2").Click
			wait(2)
            OracleListOfValues("Show Field").Select "Requester                                                                                                                                                                                                                  *REQUESTER_NAME"
            OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Requester", RACK_GetData("Vertex_Data", "Requester")
			wait(2)
			OracleFormWindow("Folder Tools").CloseWindow
            'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		   End If
	       wait(2)
  			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 1,"Amount",RACK_GetData("Vertex_Data", "Item_Amount")
		   wait(2)
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 2,"Type",RACK_GetData("Vertex_Data", "Type")
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 2,"Amount",RACK_GetData("Vertex_Data", "Tax_Amount")
		    wait(5)
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 1,"Type","Item"
             RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleButton("Distributions").Click
		wait(2)
			OracleFormWindow("Distributions").OracleTable("Table").EnterField 1,"Amount",RACK_GetData("Vertex_Data", "Item_Amount")
		wait(2)
           OracleFormWindow("Distributions").OracleTable("Table").EnterField 1,"Account",RACK_GetData("Vertex_Data","Account_Number") 
		wait(2)
		   OracleFormWindow("Distributions").SelectMenu "File->Save"
		   Verify_Oracle_Status("FRM-40400:Transaction complete: 2 records applied and saved.")
		   'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		   OracleFormWindow("Distributions").CloseWindow
		   RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		 wait(5)
				 End If
				 	If RACK_GetData("Vertex_Data","Calculate")= "Yes" Then
					 OracleFormWindow("Invoice Workbench").SelectMenu "Tools->Calculate Tax"
					 wait(2)
					 RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
					 End If
				'If OracleFormWindow("Invoice Workbench").Exist(10)  Then
				If RACK_GetData("Vertex_Data","Actions")= "Yes" Then
				OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
				wait(3)
				OracleFormWindow("Invoice Actions").OracleCheckbox("Validate").Select
				RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				wait(2)
				OracleFormWindow("Invoice Actions").OracleButton("OK").Click
				wait(2)
				OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
				wait(2)
				If RACK_GetData("Vertex_Data","Create_Accounting")= "Yes" Then
                OracleFormWindow("Invoice Actions").OracleCheckbox("Create Accounting").Select
				OracleFormWindow("Invoice Actions").OracleRadioGroup("Draft").Select "Final"
				OracleFormWindow("Invoice Actions").OracleCheckbox("Initiate Approval").Select
				End If
                If RACK_GetData("Vertex_Data","Force_Approval")= "Yes" Then
					OracleFormWindow("Invoice Actions").OracleCheckbox("Force Approval").Select
					End If
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
                OracleFormWindow("Invoice Actions").OracleButton("OK").Click
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
OracleNotification("Note").Approve
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
wait(5)
End If
End Function

'#####################################################################################################################################
'Function Description   : Function to create a Vertex Enter Mutiple nvoice Batches
'Input Parameters 	: None
'Return Value    	: None
'#####################################################################################################################################
Function Enter_Mutiple_Invoices()
            OracleFormWindow("Invoice Workbench").OracleTable("Table_2").SetFocus 1,"Type"
			OracleFormWindow("Invoice Workbench").SelectMenu "File->New"
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Type", RACK_GetData("Vertex_Data", "Invoice_Type")
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Trading Partner", RACK_GetData("Vertex_Data", "Supplier_Name")
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Invoice Date", RACK_GetData("Vertex_Data", "Invoice_Date") 
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Invoice Num", RACK_GetData("Vertex_Data", "Invoice_Number")
			Update_Notepad "Invoice_Number", Invoice_Number
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Invoice Amount", RACK_GetData("Vertex_Data", "Invoice_Amount")
			wait(2)
			If not  RACK_GetData("Vertex_Data", "Requester") = "" Then
            OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Requester", RACK_GetData("Vertex_Data", "Requester")
           wait(2)
  			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 1,"Amount",RACK_GetData("Vertex_Data", "Item_Amount")
		   wait(2)
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 2,"Type",RACK_GetData("Vertex_Data", "Type")
	       wait(2)
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 2,"Amount",RACK_GetData("Vertex_Data", "Tax_Amount")
		   wait(2)
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 1,"Type","Item"
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleButton("Distributions").Click
		   wait(2)
			OracleFormWindow("Distributions").OracleTable("Table").EnterField 1,"Amount",RACK_GetData("Vertex_Data", "Item_Amount")
		   wait(2)
           OracleFormWindow("Distributions").OracleTable("Table").EnterField 1,"Account",RACK_GetData("Vertex_Data","Account_Number") 
		   OracleFormWindow("Distributions").SelectMenu "File->Save"
		   Verify_Oracle_Status("FRM-40400:Transaction complete: 2 records applied and saved.")
		   'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		   OracleFormWindow("Distributions").CloseWindow
		   RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		    wait(5)
			If RACK_GetData("Vertex_Data","Calculate")= "Yes" Then
					 OracleFormWindow("Invoice Workbench").SelectMenu "Tools->Calculate Tax"
					 wait(2)
					 RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
					 End If
				'If OracleFormWindow("Invoice Workbench").Exist(10)  Then
				If RACK_GetData("Vertex_Data","Actions")= "Yes" Then
				OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
				wait(3)
				OracleFormWindow("Invoice Actions").OracleCheckbox("Validate").Select
				RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				wait(2)
				OracleFormWindow("Invoice Actions").OracleButton("OK").Click
				wait(2)
				OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
				wait(2)
				If RACK_GetData("Vertex_Data","Create_Accounting")= "Yes" Then
                OracleFormWindow("Invoice Actions").OracleCheckbox("Create Accounting").Select
				OracleFormWindow("Invoice Actions").OracleRadioGroup("Draft").Select "Final"
				OracleFormWindow("Invoice Actions").OracleCheckbox("Initiate Approval").Select
				End If
                If RACK_GetData("Vertex_Data","Force_Approval")= "Yes" Then
					OracleFormWindow("Invoice Actions").OracleCheckbox("Force Approval").Select
					End If
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
                OracleFormWindow("Invoice Actions").OracleButton("OK").Click
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
OracleNotification("Note").Approve
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
wait(5)
End If
End If
End Function

'#####################################################################################################################
'Function Description   : Function to RS US PYBL Process revised request set
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Vertex_RS_US_PYBL_Process()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Vertex_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Vertex_Data", "Functionality_Link"))
    If OracleFormWindow("Submit a New Request").Exist(90) Then
		OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request Set").Exist(10) Then
			OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("Vertex_Data", "Request_Name")
			wait(2)
            OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
            OracleFlexWindow("Parameters").OracleTextField("Option").Enter RACK_GetData("Vertex_Data", "Option")
			OracleFlexWindow("Parameters").OracleTextField("Invoice Batch Name").Enter RACK_GetData("Vertex_Data", "BatchName")
             OracleFlexWindow("Parameters").Approve
			 wait(2)
			 OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 2,"Parameters"
			 OracleFlexWindow("Parameters").OracleTextField("Option").Enter RACK_GetData("Vertex_Data", "Option")
			 OracleFlexWindow("Parameters").OracleTextField("Invoice Batch Name").Enter RACK_GetData("Vertex_Data", "BatchName")
             OracleFlexWindow("Parameters").Approve
			 wait(2)
            OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 3,"Parameters"
            OracleFlexWindow("Parameters").OracleTextField("Ledger").Enter RACK_GetData("Vertex_Data", "Ledger")
            OracleFlexWindow("Parameters").OracleTextField("End Date").Enter RACK_GetData("Vertex_Data", "End_Date")
			OracleFlexWindow("Parameters").OracleTextField("Report").Enter "Summary"
            OracleFlexWindow("Parameters").Approve
			wait(2)
			OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 4,"Parameters"
			OracleFlexWindow("Parameters").OracleTextField("Ledger").Enter RACK_GetData("Vertex_Data", "Ledger")
			OracleFlexWindow("Parameters").OracleTextField("End Date").Enter RACK_GetData("Vertex_Data", "End_Date")
            'OracleFlexWindow("Parameters").OracleTextField("GL Date").Enter RACK_GetData("Vertex_Data", "GL_Date")
            'OracleFlexWindow("Parameters").OracleTextField("Book").Enter RACK_GetData("Vertex_Data", "Asset_Book")
            OracleFlexWindow("Parameters").Approve
            wait(2)
			OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 5,"Parameters"
			OracleFlexWindow("Parameters").OracleTextField("GL Date").Enter RACK_GetData("Vertex_Data", "GL_Date")
            OracleFlexWindow("Parameters").OracleTextField("Book").Enter RACK_GetData("Vertex_Data", "Asset_Book")
            OracleFlexWindow("Parameters").Approve
			wait(5)
			OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 5,"Parameters"
			OracleFlexWindow("Parameters").OracleTextField("Book").Enter RACK_GetData("Vertex_Data", "Asset_Book")
			OracleFlexWindow("Parameters").Approve
			wait(2)
			OracleFormWindow("Submit Request Set").OracleButton("Submit").Click
            'RACK_ReportEvent "Validation Screenshot", "RS US PYBL Processes revised request set Parameters successfully Enter   ","Screenshot"
			RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			Handle_Oracle_Notification_Forms("OK")
			Notification_Message = Get_Oracle_Notification_Form_Message()
			Arr=split(Notification_Message, " ")
			Content = Arr(6)
			Print_Request_ID_Arr = Split(Content,")")
			Print_Request_ID = Print_Request_ID_Arr(0)
			Update_Notepad "Print_Request_ID", Print_Request_ID
			Handle_Oracle_Notification_Forms("No")
			Verify_Request_Status(Print_Request_ID)
		End if
	End if
End Function

'#####################################################################################################################
'Function Description   : Function to Vertex  Post  Journals
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Vertex_Post_Batch()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("Vertex_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("Vertex_Data", "Functionality_Link"))
		If OracleFormWindow("Find Journal Batches").Exist(120) Then
			If not RACK_GetData("Vertex_Data", "Journel_Batch_Name") = "" Then
				Query = RACK_GetData("Vertex_Data", "Journel_Batch_Name") & "%"
				OracleFormWindow("Find Journal Batches").OracleTextField("Batch").Enter Query
				OracleFormWindow("Find Journal Batches").OracleTextField("Batch").Enter Query
            	OracleFormWindow("Find Journal Batches").OracleTextField("Period").Enter RACK_GetData("Vertex_Data", "Period")
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
			wait(1)
			OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			wait(1)
			OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			wait(1)
			OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			wait(1)
			OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			wait(1)
			OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			wait(1)
			OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			wait(1)
			OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			wait(1)
			OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			wait(1)
			OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			wait(1)
			OracleFormWindow("Requests").OracleButton("Refresh Data").Click
			wait(1)
			OracleFormWindow("Requests").OracleButton("Refresh Data").Click
 Phase = OracleFormWindow("Requests").OracleTable("Table").GetFieldValue(1,"Phase")
 If Phase = "Completed" Then
 RACK_ReportEvent "Verification Post Journel -Phase is Completed","Phase = Completed","Pass"
 Else
 RACK_ReportEvent "Verification Post Journe -Phase is Completed","Phase is - "&Phase,"Fail"
			End If
		End If
End Function
'#####################################################################################################################
'Function Description   : Function to Create Purchase Order
'Input Parameters 	: None
'Return Value    	: None
'#####################################################################################################################
Function Vertex_Create_Purchase_Order()
	Select_Link(RACK_GetData("Vertex_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Vertex_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Purchase Orders").OracleTextField("Type").SetFocus
	If not  RACK_GetData("Vertex_Data", "PO_Type") = "" Then
		OracleFormWindow("Purchase Orders").OracleTextField("Type").Enter RACK_GetData("Vertex_Data", "PO_Type")
	End If
	OracleFormWindow("Purchase Orders").OracleTextField("Supplier").Enter RACK_GetData("Vertex_Data", "Supplier_Name")
	If not RACK_GetData("Vertex_Data", "Supplier_To_Site") = "" Then
		OracleFormWindow("Purchase Orders").OracleTextField("Site").Enter RACK_GetData("Vertex_Data", "Supplier_To_Site")
	End If
	If not RACK_GetData("Vertex_Data", "PO_Item") = "" Then
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").SetFocus 1,"Item"
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Item",RACK_GetData("Vertex_Data", "PO_Item")
	End If
	If not RACK_GetData("Vertex_Data", "PO_Category") = "" Then
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").SetFocus 1,"Category"
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Category",RACK_GetData("Vertex_Data", "PO_Category")
			wait(2)
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Description",RACK_GetData("Vertex_Data", "PO_Description")
	End If
	OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table").EnterField 1,"Quantity",RACK_GetData("Vertex_Data", "PO_Quantity")
	wait(2)
	OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table_3").EnterField 1,"Price",RACK_GetData("Vertex_Data", "PO_Price")
	wait(2)
	OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTable("Table_4").EnterField 1,"Need-By",RACK_GetData("Vertex_Data", "PO_Need_By_Date")
	wait(2)
	'RACK_ReportEvent "Validation Screenshot", "Purchase Orders Parameters successfully Enter  ","Screenshot"
	OracleFormWindow("Purchase Orders").OracleButton("Shipments").Click
	wait(2)
	OracleFormWindow("Shipments").OracleTabbedRegion("More").OracleTable("Table").EnterField 1,"Match  Approval Level","2-Way"
	wait(2)
	RACK_ReportEvent "Validation Screenshot", "Shipments Parameters successfully Enter  ","Screenshot"
	OracleFormWindow("Shipments").OracleButton("Distributions").Click
	wait(2)
	'OracleFormWindow("Distributions").OracleTabbedRegion("Destination").OracleTable("Table").OpenDialog 1,"Requester"
	OracleFormWindow("Distributions").OracleTabbedRegion("Destination").OracleTable("Table").EnterField 1,"Requester","CAMACHO, LORI"
'	OracleListOfValues("Requestors").Find "CAMACHO, LORI"
'    OracleListOfValues("Requestors").Select "CAMACHO, LORI"
    OracleFormWindow("Distributions").SelectMenu "File->Save"
'	If not RACK_GetData("Vertex_Data", "Po_Charge_Account") = "" Then
'		OracleFormWindow("Distributions").OracleTabbedRegion("Destination").OracleTable("Table").SetFocus 1,"PO Charge Account"
'		'OracleFlexWindow("Charge Account").Cancel
'		OracleFormWindow("Distributions").OracleTabbedRegion("Destination").OracleTable("Table_2").EnterField 1,"PO Charge Account",RACK_GetData("Vertex_Data", "Po_Charge_Account")
'		'RACK_ReportEvent "Validation Screenshot", "Distributions Parameters successfully Enter  ","Screenshot"
'	End If
'	wait(2)
'	OracleFormWindow("Distributions").SelectMenu "File->Save"
	wait(2)
	Verify_Oracle_Status("FRM-40400: Transaction complete: 4 records applied and saved.")
	OracleFormWindow("Distributions").CloseWindow
	wait(2)
	OracleFormWindow("Shipments").CloseWindow
	wait(2)	
	Purchase_Order_Number = OracleFormWindow("Purchase Orders").OracleTextField("PO, Rev").GetROProperty("value")
	RACK_PutData "Vertex_Data", "Purchase_Order_Number", CStr(Purchase_Order_Number)
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
		OracleFormWindow("Approve Document").OracleTabbedRegion("Approval Details").OracleCheckbox("Transmission Methods|E-Mail").Select
		RACK_ReportEvent "Validation Screenshot", "Approve Document Parameters successfully Enter  ","Screenshot"
		wait(2)
		OracleFormWindow("Approve Document").OracleButton("OK").Click
		wait(2)
	End if	
End Function

'#####################################################################################################################
'Function Description   : Function to create a Payment Invoice
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Vertex_Create_Po_to_Ap_Invoices()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Vertex_Data", "Responsibility_Link"))
	wait(5)
	Parent_desc.Link("Invoice Batches").click
		If OracleFormWindow("Invoice Batches").OracleTable("Table").Exist(120) Then
            OracleFormWindow("Invoice Batches").OracleTable("Table").EnterField 1,"Batch Name", RACK_GetData("Vertex_Data", "BatchName")
			RACK_ReportEvent "Validation Screenshot", "Po BatchName successfully Enter ","Screenshot"
			wait(2)
			OracleFormWindow("Invoice Batches").OracleButton("Invoices").Click
			wait(5)
			If OracleFormWindow("Invoice Workbench").OracleTable("Table").Exist(20) Then
				If not RACK_GetData("Vertex_Data", "Purchase_Order_Number") = "" Then
					PO_Number = RACK_GetData("Vertex_Data", "Purchase_Order_Number")
				else
					PO_Number = Read_Notepad("PO_Number")
				End if
                OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"PO Number", PO_Number
				wait(2)
				OracleFormWindow("Invoice Workbench").OracleTable("Table").SetFocus 1,"Invoice Date"                     
				OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Date",RACK_GetData("Vertex_Data", "Invoice_Date")
				wait(2)    
				If not RACK_GetData("Vertex_Data", "Supplier_Name") = "" Then
				OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Trading Partner", RACK_GetData("Vertex_Data", "Supplier_Name")   
				End if
				OracleFormWindow("Invoice Workbench").OracleTable("Table").SetFocus 1,"Invoice Num" 
				OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Num", RACK_GetData("Vertex_Data", "Invoice_Number")
                 wait(2)
				'OracleFormWindow("Invoice Workbench").OracleTable("Table").SetFocus 1,"Invoice Amount"
				OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Amount", RACK_GetData("Vertex_Data", "Invoice_Amount")
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				wait(2)
				OracleFormWindow("Invoice Workbench").OracleButton("Match").Click
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				'RACK_ReportEvent "Validation Screenshot", "Po Match Button successfully click ","Screenshot"
				If  OracleFormWindow("Find Purchase Orders for").Exist(10) Then
					OracleFormWindow("Find Purchase Orders for").OracleButton("Find").Click
				End if
				If  OracleFormWindow("Match to Purchase Orders").Exist(5) Then
                    OracleFormWindow("Match to Purchase Orders").OracleTable("Table").EnterField 1,"Match", true
				RACK_ReportEvent "Validation Screenshot", "Po Match checkbox successfully select ","Screenshot"
				wait(5)
				OracleFormWindow("Match to Purchase Orders").OracleButton("Match").Click
				wait(2)
'				OracleFormWindow("Invoice Workbench").SelectMenu "File->Save"
'                Verify_Oracle_Status("FRM-40400:Transaction complete: 2 records applied and saved.")
				End If
				If not  RACK_GetData("Vertex_Data", "Type") = "" Then
			   OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 2,"Type",RACK_GetData("Vertex_Data", "Type")
			   End If
			   If not  RACK_GetData("Vertex_Data", "Tax_Amount") = "" Then
			   OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 2,"Amount",RACK_GetData("Vertex_Data", "Tax_Amount")
			   OracleFormWindow("Invoice Workbench").SelectMenu "File->Save"
               Verify_Oracle_Status("FRM-40400:Transaction complete: 2 records applied and saved.")
			   End If
			   If RACK_GetData("Vertex_Data","Calculate")= "Yes" Then
					 OracleFormWindow("Invoice Workbench").SelectMenu "Tools->Calculate Tax"
					 wait(2)
					 RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
					 End If
				If RACK_GetData("Vertex_Data","Actions")= "Yes" Then
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
End If
End Function
'#####################################################################################################################
'Function Description   : Function to create a Payment Invoice
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Vertex_Find_Po_to_Ap_Invoices()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Vertex_Data", "Responsibility_Link"))
	wait(5)
	Parent_desc.Link("Invoice Batches").click
	wait(60)
	    OracleFormWindow("Invoice Batches").SelectMenu "View->Find..."
		wait(2)
		OracleFormWindow("Find Invoice Batches").OracleTextField("Names").Enter RACK_GetData("Vertex_Data", "BatchName")
		wait(2)
		OracleFormWindow("Find Invoice Batches").OracleButton("Find").Click
		wait(2)
		OracleFormWindow("Invoice Batches").OracleButton("Invoices").Click
		wait(2)
		OracleFormWindow("Invoice Workbench").OracleTable("Table_2").SetFocus 1,"Type"
		OracleFormWindow("Invoice Workbench").SelectMenu "File->New"
			If OracleFormWindow("Invoice Workbench").OracleTable("Table").Exist(20) Then
				If not RACK_GetData("Vertex_Data", "Purchase_Order_Number") = "" Then
					PO_Number = RACK_GetData("Vertex_Data", "Purchase_Order_Number")
				else
					PO_Number = Read_Notepad("PO_Number")
				End if
                OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"PO Number", PO_Number
				wait(2)
				OracleFormWindow("Invoice Workbench").OracleTable("Table_2").SetFocus 2,"Invoice Date"                     
				OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Invoice Date",RACK_GetData("Vertex_Data", "Invoice_Date")
				wait(2)    
				If not RACK_GetData("Vertex_Data", "Supplier_Name") = "" Then
				OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Trading Partner", RACK_GetData("Vertex_Data", "Supplier_Name")   
				End if
				OracleFormWindow("Invoice Workbench").OracleTable("Table_2").SetFocus 2,"Invoice Num" 
				OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Invoice Num", RACK_GetData("Vertex_Data", "Invoice_Number")
                 wait(2)
				'OracleFormWindow("Invoice Workbench").OracleTable("Table").SetFocus 1,"Invoice Amount"
				OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Invoice Amount", RACK_GetData("Vertex_Data", "Invoice_Amount")
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				wait(2)
				OracleFormWindow("Invoice Workbench").OracleButton("Match").Click
				'RACK_ReportEvent "Validation Screenshot", "Po Match Button successfully click ","Screenshot"
				If  OracleFormWindow("Find Purchase Orders for").Exist(10) Then
					OracleFormWindow("Find Purchase Orders for").OracleButton("Find").Click
				End if
				If  OracleFormWindow("Match to Purchase Orders").Exist(5) Then
                    OracleFormWindow("Match to Purchase Orders").OracleTable("Table").EnterField 1,"Match", true
				RACK_ReportEvent "Validation Screenshot", "Po Match checkbox successfully select ","Screenshot"
				wait(5)
				OracleFormWindow("Match to Purchase Orders").OracleButton("Match").Click
				wait(2)
'				OracleFormWindow("Invoice Workbench").SelectMenu "File->Save"
'                Verify_Oracle_Status("FRM-40400:Transaction complete: 2 records applied and saved.")
				End If
				If not  RACK_GetData("Vertex_Data", "Type") = "" Then
			   OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 2,"Type",RACK_GetData("Vertex_Data", "Type")
			   End If
			   If not  RACK_GetData("Vertex_Data", "Tax_Amount") = "" Then
			   OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 2,"Amount",RACK_GetData("Vertex_Data", "Tax_Amount")
			   OracleFormWindow("Invoice Workbench").SelectMenu "File->Save"
               Verify_Oracle_Status("FRM-40400:Transaction complete: 2 records applied and saved.")
			   End If
			   If RACK_GetData("Vertex_Data","Calculate")= "Yes" Then
					 OracleFormWindow("Invoice Workbench").SelectMenu "Tools->Calculate Tax"
					 wait(2)
					 RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
					 End If
				If RACK_GetData("Vertex_Data","Actions")= "Yes" Then
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
'Function Description   : Function to create a Payment Invoice
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Enter_Mutiple_Po_to_Ap_Invoices()
            OracleFormWindow("Invoice Workbench").OracleTable("Table_2").SetFocus 1,"Type"
			OracleFormWindow("Invoice Workbench").SelectMenu "File->New"
			wait(5)
			If OracleFormWindow("Invoice Workbench").OracleTable("Table").Exist(20) Then
				If not RACK_GetData("Vertex_Data", "Purchase_Order_Number") = "" Then
					PO_Number = RACK_GetData("Vertex_Data", "Purchase_Order_Number")
				else
					PO_Number = Read_Notepad("PO_Number")
				End if
                OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"PO Number", PO_Number
				wait(2)
				OracleFormWindow("Invoice Workbench").OracleTable("Table_2").SetFocus 2,"Invoice Date"                     
				OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Invoice Date",RACK_GetData("Vertex_Data", "Invoice_Date")
				wait(2)    
				If not RACK_GetData("Vertex_Data", "Supplier_Name") = "" Then
				OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Trading Partner", RACK_GetData("Vertex_Data", "Supplier_Name")   
				End if
				OracleFormWindow("Invoice Workbench").OracleTable("Table_2").SetFocus 2,"Invoice Num" 
				OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 2,"Invoice Num", RACK_GetData("Vertex_Data", "Invoice_Number")
                 wait(2)
				'OracleFormWindow("Invoice Workbench").OracleTable("Table").SetFocus 1,"Invoice Amount"
				OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 2,"Invoice Amount", RACK_GetData("Vertex_Data", "Invoice_Amount")
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				wait(2)
				OracleFormWindow("Invoice Workbench").OracleButton("Match").Click
				'RACK_ReportEvent "Validation Screenshot", "Po Match Button successfully click ","Screenshot"
				If  OracleFormWindow("Find Purchase Orders for").Exist(10) Then
					OracleFormWindow("Find Purchase Orders for").OracleButton("Find").Click
				End if
				If  OracleFormWindow("Match to Purchase Orders").Exist(5) Then
                    OracleFormWindow("Match to Purchase Orders").OracleTable("Table").EnterField 1,"Match", true
				RACK_ReportEvent "Validation Screenshot", "Po Match checkbox successfully select ","Screenshot"
				wait(5)
				OracleFormWindow("Match to Purchase Orders").OracleButton("Match").Click
				wait(2)
'				OracleFormWindow("Invoice Workbench").SelectMenu "File->Save"
'                Verify_Oracle_Status("FRM-40400:Transaction complete: 2 records applied and saved.")
				End If
				If not  RACK_GetData("Vertex_Data", "Type") = "" Then
			   OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 2,"Type",RACK_GetData("Vertex_Data", "Type")
			   End If
			   If not  RACK_GetData("Vertex_Data", "Tax_Amount") = "" Then
			   OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 2,"Amount",RACK_GetData("Vertex_Data", "Tax_Amount")
			   OracleFormWindow("Invoice Workbench").SelectMenu "File->Save"
               Verify_Oracle_Status("FRM-40400:Transaction complete: 2 records applied and saved.")
			   End If
			   If RACK_GetData("Vertex_Data","Calculate")= "Yes" Then
					 OracleFormWindow("Invoice Workbench").SelectMenu "Tools->Calculate Tax"
					 wait(2)
					 RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
					 End If
				If RACK_GetData("Vertex_Data","Actions")= "Yes" Then
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
'''#####################################################################################################################
'''Function Description   : Function for Vertex AP Invoice Validation - Pre Process
'''Input Parameters 	: None
'''Return Value    	: None
'''##################################################################################################################### 
'Function Vertex_AP_Invoice_Request_Set()
'	set WshShell = CreateObject("WScript.Shell")
'	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
'	Select_Link(RACK_GetData("Vertex_Data", "Responsibility_Link"))
'	wait(5)
'	Select_Link(RACK_GetData("Vertex_Data", "Functionality_Link"))
'    If OracleFormWindow("Submit a New Request").Exist(90) Then
'		OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
'		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
'		If OracleFormWindow("Submit Request Set").Exist(10) Then
'			OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("Vertex_Data", "Request_Name")
'			wait(2)
'			OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
'			wait(2)
'			'If not  RACK_GetData("Vertex_Data", "Operating_Unit") = "" Then
'            'OracleFlexWindow("Parameters").OracleTextField("Operating Unit").Enter RACK_GetData("Vertex_Data", "Operating_Unit")
'            OracleFlexWindow("Parameters").OracleTextField("Option").Enter RACK_GetData("Vertex_Data", "Option")
'			wait(2)
'            OracleFlexWindow("Parameters").OracleTextField("Invoice Batch Name").Enter RACK_GetData("Vertex_Data", "BatchName")
'            OracleFlexWindow("Parameters").Approve
'			wait(5)
'            OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 2,"Parameters"
'            OracleFlexWindow("Parameters").OracleTextField("Invoice Batch Name").Enter RACK_GetData("Vertex_Data", "BatchName")
'            OracleFlexWindow("Parameters").Approve
'			wait(2)
'			OracleFormWindow("Submit Request Set").OracleButton("Submit").Click
'			RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
'			Notification_Message = Get_Oracle_Notification_Form_Message()
'			Arr=split(Notification_Message, " ")
'			Content = Arr(6)
'			Print_Request_ID_Arr = Split(Content,")")
'			Print_Request_ID = Print_Request_ID_Arr(0)
'			Update_Notepad "Print_Request_ID", Print_Request_ID
'			Handle_Oracle_Notification_Forms("No")
'			Verify_Request_Status(Print_Request_ID)
'		else
'			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
'		end if
'	    else
'			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
'	End If
'End Function
''#####################################################################################################################
''Function Description   : Function to Create Accounting
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 
'Function Vertex_Create_Accounting()
'	set WshShell = CreateObject("WScript.Shell")
'	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
'	Select_Link(RACK_GetData("Vertex_Data", "Responsibility_Link"))
'	wait(3)
'	Select_Link(RACK_GetData("Vertex_Data", "Functionality_Link"))
'    If OracleFormWindow("Submit a New Request").Exist(90) Then
'		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
'		If OracleFormWindow("Submit Request").Exist(10) Then
'			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Vertex_Data", "Request_Name")
'			wait(2)
'              OracleFlexWindow("Parameters").OracleTextField("Ledger").Enter RACK_GetData("Vertex_Data", "Ledger")
'              OracleFlexWindow("Parameters").OracleTextField("End Date").Enter RACK_GetData("Vertex_Data", "End_Date")
'              OracleFlexWindow("Parameters").OracleTextField("Report").Enter RACK_GetData("Vertex_Data", "Report")
'              OracleFlexWindow("Parameters").Approve
'			wait(2)
'			OracleFormWindow("Submit Request").OracleButton("Submit").Click
'			RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
'			wait(5)
'			Notification_Message = Get_Oracle_Notification_Form_Message()
'			Arr=split(Notification_Message, " ")
'			Content = Arr(6)
'			Print_Request_ID_Arr = Split(Content,")")
'			Print_Request_ID = Print_Request_ID_Arr(0)
'			Update_Notepad "Print_Request_ID", Print_Request_ID
'			Handle_Oracle_Notification_Forms("No")
'			Verify_Request_Status(Print_Request_ID)
'		else
'			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
'		end if
'	    else
'			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
'	End If
'End Function
''#####################################################################################################################
''Function Description   : Function to Vertex Transfer JournalEntries to GL
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 
'Function Vertex_Transfer_to_GL()
'	set WshShell = CreateObject("WScript.Shell")
'	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
'	Select_Link(RACK_GetData("Vertex_Data", "Responsibility_Link"))
'	wait(3)
'	Select_Link(RACK_GetData("Vertex_Data", "Functionality_Link"))
'	If OracleFormWindow("Submit a New Request").Exist(90) Then
'		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
'		If OracleFormWindow("Submit Request").Exist(10) Then
'			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Vertex_Data", "Request_Name")
'			wait(2)
'			If  OracleFlexWindow("Parameters").Exist(10) Then
'				OracleFlexWindow("Parameters").OracleTextField("Ledger").Enter RACK_GetData("Vertex_Data", "Ledger")
'				OracleFlexWindow("Parameters").OracleTextField("End Date").Enter RACK_GetData("Vertex_Data", "End_Date")
'				OracleFlexWindow("Parameters").OracleButton("OK").Click
'				wait(2)
'			End If
'			OracleFormWindow("Submit Request").OracleButton("Submit").Click
'            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
'		End If
'		Notification_Message = Get_Oracle_Notification_Form_Message()
'		Arr=split(Notification_Message, " ")
'		Content = Arr(6)
'		Payment_Request_ID_Arr = Split(Content,")")
'		Payment_Request_ID = Payment_Request_ID_Arr(0)
'		Update_Notepad "Payment_Request_ID", Payment_Request_ID
'		Handle_Oracle_Notification_Forms("No")
'	End If
'	Verify_Request_Status(Payment_Request_ID)
'End Function

'#####################################################################################################################
'Function Description   : Function to create a Invoice Batches
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Create_Manual_Invoice()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("AP_Data", "Responsibility_Link"))
	wait(5)
	Parent_desc.Link("Invoice Batches").click
		If OracleFormWindow("Invoice Batches").OracleTable("Table").Exist(120) Then
        OracleFormWindow("Invoice Batches").OracleTable("Table").EnterField 1,"Batch Name", RACK_GetData("AP_Data", "BatchName")
        RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(5)
		OracleFormWindow("Invoice Batches").OracleButton("Invoices").Click
		wait(2)
		End If
            If OracleFormWindow("Invoice Workbench").OracleTable("Table").Exist(10) Then
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Type", RACK_GetData("AP_Data", "Invoice_Type")
			wait(5)
			'OracleNotification("Note").Approve
			'OracleFormWindow("Invoice Workbench").OracleTextField("Trading Partner").Enter RACK_GetData("AP_Data","Trading_Partner")  
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Trading Partner", RACK_GetData("AP_Data", "Supplier_Name")
			'OracleFormWindow("Invoice Workbench").OracleTextField("Supplier Num").Click
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Date", RACK_GetData("AP_Data", "Invoice_Date")
			'OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Date").Enter RACK_GetData("AP_Data","Invoice_Date")   
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Num", RACK_GetData("AP_Data", "Invoice_Number")
			Update_Notepad "Invoice_Number", Invoice_Number
		'	OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Num").Enter RACK_GetData("AP_Data","Invoice_Number") 
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Amount", RACK_GetData("AP_Data", "Invoice_Amount")
			'OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Amount").Enter RACK_GetData("AP_Data","Invoice_Amount")  
			If not  RACK_GetData("AP_Data", "Requester") = "" Then
			OracleFormWindow("Invoice Workbench").SelectMenu "Folder->Folder Tools"
			wait(2)
            OracleFormWindow("Folder Tools").OracleButton("Show Field...").Click
			wait(2)
            OracleListOfValues("Show Field").Select "Requester                                                                                                                                                                                                                  *REQUESTER_NAME"
            OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 1,"Requester", RACK_GetData("AP_Data", "Requester")
			wait(2)
			'OracleFormWindow("Invoice Workbench").OracleTable("Table_2").OpenDialog 1,"Requester"
            'OracleListOfValues("Requester").Select "CAMACHO, LORI"
			OracleFormWindow("Folder Tools").CloseWindow
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		  End If
           If not  RACK_GetData("AP_Data", "Line_Type") = "" Then
           OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 1,"Type",RACK_GetData("AP_Data","Line_Type")
		   End If
	       wait(2)
           If not  RACK_GetData("AP_Data", "Amount") = "" Then
  			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 1,"Amount",RACK_GetData("AP_Data","Amount")
			End If
		  'OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Amount").Enter RACK_GetData("AP_Data","Amount")   
'			wait(2)
'		   If not  RACK_GetData("AP_Data", "Tax_Regime") = "" Then
'		   OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").OpenDialog 1,"Tax Regime",RACK_GetData("AP_Data","Tax_Regime")
'		   end If
'		   wait(2)
'		    If not  RACK_GetData("AP_Data", "Tax") = "" Then
'           OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").OpenDialog 1,"Tax",RACK_GetData("AP_Data","Tax")
'		   end If
'		   wait(2)
'		    If not  RACK_GetData("AP_Data", "Tax_Status") = "" Then
'           OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").OpenDialog 1,"Tax Status",RACK_GetData("AP_Data","Tax_Status")
'		   end If
'		   wait(2)
'		    If not  RACK_GetData("AP_Data", "Tax_Rate_Name") = "" Then
'           OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").OpenDialog 1,"Tax Rate Name",RACK_GetData("AP_Data","Tax_Rate_Name")
'		   end If
'           RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		   wait(2)
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleButton("Distributions").Click
		   wait(2)
			End If
			   If OracleFormWindow("Distributions").Exist(60)  Then
				OracleFormWindow("Distributions").OracleTable("Table").EnterField 1,"Amount",RACK_GetData("AP_Data","Amount")
			'	OracleFormWindow("Distributions").OracleTextField("Amount").Enter RACK_GetData("AP_Data","Amount_1")   
				wait(2)
				'OracleFormWindow("Distributions").OracleTable("Table").SetFocus 1,"Account"
				'OracleFormWindow("Distributions").OracleTextField("Account").Click
				wait(2)
				OracleFormWindow("Distributions").OracleTable("Table").EnterField 1,"Account",RACK_GetData("AP_Data","Account_Number") 
				wait(2)
				OracleFormWindow("Distributions").SelectMenu "File->Save"
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				wait(2)
				End If
				If RACK_GetData("AP_Data","Actions")= "Yes" Then
				OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
				wait(3)
				OracleFormWindow("Invoice Actions").OracleCheckbox("Validate").Select
				RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				wait(2)
				OracleFormWindow("Invoice Actions").OracleButton("OK").Click
				wait(2)
				OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
				wait(2)
				If RACK_GetData("AP_Data","Create_Accounting")= "Yes" Then
                OracleFormWindow("Invoice Actions").OracleCheckbox("Create Accounting").Select
				OracleFormWindow("Invoice Actions").OracleRadioGroup("Draft").Select "Final"
				OracleFormWindow("Invoice Actions").OracleCheckbox("Initiate Approval").Select
				End If
                If RACK_GetData("AP_Data","Force_Approval")= "Yes" Then
					OracleFormWindow("Invoice Actions").OracleCheckbox("Force Approval").Select
					End If
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
                OracleFormWindow("Invoice Actions").OracleButton("OK").Click
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
OracleNotification("Note").Approve
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
End If
End Function
'#####################################################################################################################
'Function Description   : Function to create a Payment Invoice
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Match_PO()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("AP_Data", "Responsibility_Link"))
	wait(5)
	Parent_desc.Link("Invoice Batches").click
		If OracleFormWindow("Invoice Batches").OracleTable("Table").Exist(120) Then
            OracleFormWindow("Invoice Batches").OracleTable("Table").EnterField 1,"Batch Name", RACK_GetData("AP_Data", "BatchName")
			RACK_ReportEvent "Validation Screenshot", "Po BatchName successfully Enter ","Screenshot"
			wait(2)
			OracleFormWindow("Invoice Batches").OracleButton("Invoices").Click
			wait(5)
			If OracleFormWindow("Invoice Workbench").OracleTable("Table").Exist(20) Then
				If not RACK_GetData("AP_Data", "Purchase_Order_Number") = "" Then
					PO_Number = RACK_GetData("AP_Data", "Purchase_Order_Number")
				else
					PO_Number = Read_Notepad("PO_Number")
				End if
                OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"PO Number", PO_Number
				wait(2)
				OracleFormWindow("Invoice Workbench").OracleTable("Table").SetFocus 1,"Invoice Date"                     
				OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Date",RACK_GetData("AP_Data", "Invoice_Date")
				wait(2)    
				If not RACK_GetData("AP_Data", "Supplier_Name") = "" Then
				OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Trading Partner", RACK_GetData("AP_Data", "Supplier_Name")   
				End if
				OracleFormWindow("Invoice Workbench").OracleTable("Table").SetFocus 1,"Invoice Num" 
				wait(5)
				Handle_Oracle_Notification_Forms("Cancel")
				OracleFormWindow("Invoice Workbench").OracleTable("Table").SetFocus 1,"Invoice Num" 
				Invoice_Number = RACK_GetData("AP_Data", "Invoice_Number")
				OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Num", Invoice_Number
				Update_Notepad "Invoice_Number", Invoice_Number
				OracleFormWindow("Invoice Workbench").OracleTable("Table").SetFocus 1,"Invoice Amount"
				OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Amount", RACK_GetData("AP_Data", "Invoice_Amount")
'				RACK_ReportEvent "Validation Screenshot", "Invoice Workbench Parameters successfully Enter   ","Screenshot"
				wait(2)
				If not RACK_GetData("AP_Data", "Requester") = "" Then
				OracleFormWindow("Invoice Workbench").SelectMenu "Folder->Folder Tools"
				wait(2)
                OracleFormWindow("Folder Tools").OracleButton("Show Field...").Click
				wait(2)
                OracleListOfValues("Show Field").Select "Requester                                                                                                                                                                                                                  *REQUESTER_NAME"
                OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 1,"Requester",RACK_GetData("AP_Data", "Requester")
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
				If  OracleFormWindow("Match to Purchase Orders").Exist(5) Then
                    OracleFormWindow("Match to Purchase Orders").OracleTable("Table").EnterField 1,"Match", true
				RACK_ReportEvent "Validation Screenshot", "Po Match checkbox successfully select ","Screenshot"
				wait(2)
				OracleFormWindow("Match to Purchase Orders").OracleButton("Match").Click
				wait(5)
                Verify_Oracle_Status("FRM-40400:Transaction complete: 1 records applied and saved.")
				End If
				End If
				If  RACK_GetData("AP_Data", "Invoice_Actions") = "Yes" Then
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
'Function Description   : Function to create a Invoice Batches
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Create_Credit_Memo()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("AP_Data", "Responsibility_Link"))
	wait(5)
	Parent_desc.Link("Invoice Batches").click
		If OracleFormWindow("Invoice Batches").OracleTable("Table").Exist(120) Then
        OracleFormWindow("Invoice Batches").OracleTable("Table").EnterField 1,"Batch Name", RACK_GetData("AP_Data", "BatchName")
        RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(2)
		OracleFormWindow("Invoice Batches").OracleButton("Invoices").Click
		wait(5)
		End If
            If OracleFormWindow("Invoice Workbench").OracleTable("Table").Exist(10) Then
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Type", RACK_GetData("AP_Data", "Invoice_Type")
			wait(2)
			'OracleNotification("Note").Approve
			'OracleFormWindow("Invoice Workbench").OracleTextField("Trading Partner").Enter RACK_GetData("AP_Data","Trading_Partner")  
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Trading Partner", RACK_GetData("AP_Data", "Supplier_Name")
			'OracleFormWindow("Invoice Workbench").OracleTextField("Supplier Num").Click
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Date", RACK_GetData("AP_Data", "Invoice_Date")
			'OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Date").Enter RACK_GetData("AP_Data","Invoice_Date")   
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Num", RACK_GetData("AP_Data", "Invoice_Number")
			Update_Notepad "Invoice_Number", Invoice_Number
		'	OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Num").Enter RACK_GetData("AP_Data","Invoice_Number") 
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Amount", RACK_GetData("AP_Data", "Invoice_Amount")
			'OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Amount").Enter RACK_GetData("AP_Data","Invoice_Amount")  
			If not  RACK_GetData("AP_Data", "Requester") = "" Then
			OracleFormWindow("Invoice Workbench").SelectMenu "Folder->Folder Tools"
			wait(2)
            OracleFormWindow("Folder Tools").OracleButton("Show Field...").Click
			wait(2)
            OracleListOfValues("Show Field").Select "Requester                                                                                                                                                                                                                  *REQUESTER_NAME"
            OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 1,"Requester", RACK_GetData("AP_Data", "Requester")
			wait(2)
			'OracleFormWindow("Invoice Workbench").OracleTable("Table_2").OpenDialog 1,"Requester"
            'OracleListOfValues("Requester").Select "CAMACHO, LORI"
			OracleFormWindow("Folder Tools").CloseWindow
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		  End If
           If not  RACK_GetData("AP_Data", "Line_Type") = "" Then
           OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 1,"Type",RACK_GetData("AP_Data","Line_Type")
		   End If
	       wait(2)
           If not  RACK_GetData("AP_Data", "Amount") = "" Then
  			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 1,"Amount",RACK_GetData("AP_Data","Amount")
			End If
		  'OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Amount").Enter RACK_GetData("AP_Data","Amount")   
'			wait(2)
'		   If not  RACK_GetData("AP_Data", "Tax_Regime") = "" Then
'		   OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").OpenDialog 1,"Tax Regime",RACK_GetData("AP_Data","Tax_Regime")
'		   end If
'		   wait(2)
'		    If not  RACK_GetData("AP_Data", "Tax") = "" Then
'           OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").OpenDialog 1,"Tax",RACK_GetData("AP_Data","Tax")
'		   end If
'		   wait(2)
'		    If not  RACK_GetData("AP_Data", "Tax_Status") = "" Then
'           OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").OpenDialog 1,"Tax Status",RACK_GetData("AP_Data","Tax_Status")
'		   end If
'		   wait(2)
'		    If not  RACK_GetData("AP_Data", "Tax_Rate_Name") = "" Then
'           OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").OpenDialog 1,"Tax Rate Name",RACK_GetData("AP_Data","Tax_Rate_Name")
'		   end If
'           RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		   wait(2)
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleButton("Distributions").Click
		   wait(2)
			End If
			   If OracleFormWindow("Distributions").Exist(60)  Then
				OracleFormWindow("Distributions").OracleTable("Table").EnterField 1,"Amount",RACK_GetData("AP_Data","Amount")
			'	OracleFormWindow("Distributions").OracleTextField("Amount").Enter RACK_GetData("AP_Data","Amount_1")   
				wait(2)
				'OracleFormWindow("Distributions").OracleTable("Table").SetFocus 1,"Account"
				'OracleFormWindow("Distributions").OracleTextField("Account").Click
				wait(2)
				OracleFormWindow("Distributions").OracleTable("Table").EnterField 1,"Account",RACK_GetData("AP_Data","Account_Number") 
				wait(2)
				OracleFormWindow("Distributions").SelectMenu "File->Save"
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				wait(2)
				End If
				If  RACK_GetData("AP_Data", "Invoice_Actions") = "Yes" Then
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
			If OracleNotification("Note").Exist(20) Then
				OracleNotification("Note").Approve
	     		Verify_Oracle_Status("FRM-40400:Transaction complete: 1 records applied and saved.")
End If
End If
End Function
'##################################################################################################################### 
'VERTEX TESTING
'Function Description   : Function to Vertex create a Invoice Batches
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Lines_Tax_Allocations()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("AP_Data", "Responsibility_Link"))
	wait(5)
	Parent_desc.Link("Invoice Batches").click
	If OracleFormWindow("Invoice Batches").OracleTable("Table").Exist(120) Then
        OracleFormWindow("Invoice Batches").OracleTable("Table").EnterField 1,"Batch Name", RACK_GetData("AP_Data", "BatchName")
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(2)
		OracleFormWindow("Invoice Batches").OracleButton("Invoices").Click
		End If
            If OracleFormWindow("Invoice Workbench").OracleTable("Table").Exist(10) Then
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Type", RACK_GetData("AP_Data", "Invoice_Type")
			wait(5)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Trading Partner", RACK_GetData("AP_Data", "Supplier_Name")
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Date", RACK_GetData("AP_Data", "Invoice_Date") 
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Num", RACK_GetData("AP_Data", "Invoice_Number")
			Update_Notepad "Invoice_Number", Invoice_Number
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTable("Table").EnterField 1,"Invoice Amount", RACK_GetData("AP_Data", "Invoice_Amount")
			wait(2)
			If not  RACK_GetData("AP_Data", "Requester") = "" Then
			OracleFormWindow("Invoice Workbench").SelectMenu "Folder->Folder Tools"
			wait(2)
            OracleFormWindow("Folder Tools").OracleButton("Show Field...").Click
			wait(2)
            OracleListOfValues("Show Field").Select "Requester                                                                                                                                                                                                                  *REQUESTER_NAME"
            OracleFormWindow("Invoice Workbench").OracleTable("Table_2").EnterField 1,"Requester", RACK_GetData("AP_Data", "Requester")
			wait(2)
			OracleFormWindow("Folder Tools").CloseWindow
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		   End If
	       wait(2)
  			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 1,"Amount",RACK_GetData("AP_Data", "Item_Amount")
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 2,"Type",RACK_GetData("AP_Data", "Type")
			wait(2)
			OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 2,"Amount",RACK_GetData("AP_Data", "Tax_Amount")
			wait(2)
		   OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTable("Table").EnterField 1,"Type","Item"
		   wait(2) 
           RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		   OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleButton("Distributions").Click
		   wait(2)
		   OracleFormWindow("Distributions").OracleTable("Table").EnterField 1,"Amount",RACK_GetData("AP_Data", "Item_Amount")
		   wait(2)
           OracleFormWindow("Distributions").OracleTable("Table").EnterField 1,"Account",RACK_GetData("AP_Data","Account_Number") 
		   wait(2)
		   OracleFormWindow("Distributions").SelectMenu "File->Save"
		   OracleFormWindow("Distributions").CloseWindow
		   RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		    End If
			If  RACK_GetData("AP_Data", "Invoice_Actions") = "Yes" Then
				OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
				wait(5)
				OracleFormWindow("Invoice Actions").OracleCheckbox("Validate").Select
				wait(2)
				OracleFormWindow("Invoice Actions").OracleButton("OK").Click
				wait(2)
				OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
				wait(2)
                OracleFormWindow("Invoice Actions").OracleCheckbox("Create Accounting").Select
				OracleFormWindow("Invoice Actions").OracleRadioGroup("Draft").Select "Final"
				OracleFormWindow("Invoice Actions").OracleCheckbox("Initiate Approval").Select
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
                OracleFormWindow("Invoice Actions").OracleButton("OK").Click
                Verify_Oracle_Status("FRM-40400:Transaction complete: 1 records applied and saved.")
			If OracleNotification("Note").Exist(20) Then
				OracleNotification("Note").Approve
	     		Verify_Oracle_Status("FRM-40400:Transaction complete: 1 records applied and saved.")
End If
wait(5)
End If
End Function
''#####################################################################################################################
''Function Description   : Function for Solution_Partner Process
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 
Function Solution_Partner()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("AP_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("AP_Data", "Functionality_Link"))
	If Parent_desc.WebFile("uploadFile_oafileUpload").Exist(20) Then
        wait(2)
		Parent_desc.WebFile("uploadFile_oafileUpload").Set RACK_GetData("AP_Data", "Attachement_Path") '"C:\Users\221045\Desktop\sample.txt"
		File_Name = Split(RACK_GetData("AP_Data", "Attachement_Path"),"\")
		Link_Name = File_Name(ubound(File_Name))
		wait(5)
		Parent_desc.WebButton("Submit").Click
		wait(5)
		If Parent_desc.Link("name:="& Link_Name).Exist(20) then
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is uploaded as expected","Pass"
		else
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is not uploaded as expected","Fail"
		end if
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	end if
End Function
'#####################################################################################################################
'Function Description   : Function to Validate and check Invoice open interface
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Payables_Open_Interface_Set()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("AP_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("AP_Data", "Functionality_Link"))
    If OracleFormWindow("Submit a New Request").Exist(90) Then
		OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request Set").Exist(10) Then
			OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("AP_Data", "Request_Name")
			wait(2)
             OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
			 If not  RACK_GetData("AP_Data", "File_Name") = "" Then
			 'OracleFlexWindow("Parameters").OracleTextField("Input File Name").Enter RACK_GetData("AP_Data", "File_Name")
             OracleFlexWindow("Parameters").OracleTextField("File Name").Enter RACK_GetData("AP_Data", "File_Name")
			 End If
             OracleFlexWindow("Parameters").Approve
			 wait(2)
             OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 2,"Parameters"
			 wait(2)
             OracleFlexWindow("Parameters").OracleTextField("Source").Enter RACK_GetData("AP_Data", "Source")
             wait(2)
             OracleFlexWindow("Parameters").OracleTextField("Batch Name").Enter RACK_GetData("AP_Data", "BatchName")
             OracleFlexWindow("Parameters").Approve
			 wait(2)
             OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 3,"Parameters"
             OracleFlexWindow("Parameters").Approve
			'End If
			wait(2)
			OracleFormWindow("Submit Request Set").OracleButton("Submit").Click
			RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			Handle_Oracle_Notification_Forms("OK")
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
		 End If
	    else
			RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	 End If
End Function
'#####################################################################################################################
'Function Description   : Function to AP and PO Accrual Reconciliation Report
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function AP_PO_Accrual_Process()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("AP_Data", "Responsibility_Link"))
	wait(3)
	'Select_Link(RACK_GetData("AP_Data", "Functionality_Link"))
	Parent_desc.Link("Run").Click
	If OracleFormWindow("Submit a New Request").Exist(90) Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(10) Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("AP_Data", "Request_Name")
			wait(2)
			If OracleFlexWindow("Parameters").Exist(4)  Then
				OracleFlexWindow("Parameters").OracleButton("OK").Click
				wait(5)
			End If
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
			RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			Handle_Oracle_Notification_Forms("OK")
			wait(5)
			Notification_Message = Get_Oracle_Notification_Form_Message()
			Arr=split(Notification_Message, " ")
			Content = Arr(6)
			Print_Request_ID_Arr = Split(Content,")")
			Print_Request_ID = Print_Request_ID_Arr(0)
			Update_Notepad "Print_Request_ID", Print_Request_ID
			Handle_Oracle_Notification_Forms("No")
			Verify_Request_Status(Print_Request_ID)
		End If
	End If
End Function
'#####################################################################################################################
'Function Description   : Function to Unaccounted Transactions Report
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Unaccounted_Transactions_Report()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("AP_Data", "Responsibility_Link"))
	wait(3)
	'Select_Link(RACK_GetData("AP_Data", "Functionality_Link"))
	Parent_desc.Link("Run").Click
	If OracleFormWindow("Submit a New Request").Exist(90) Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(10) Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("AP_Data", "Request_Name")
			wait(2)
			If OracleFlexWindow("Parameters").Exist(4)  Then
				OracleFlexWindow("Parameters").OracleButton("OK").Click
				wait(5)
			end if
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
			RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			Handle_Oracle_Notification_Forms("OK")
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
'#####################################################################################################################
'Function Description   : Function to run the request Rackspace Key Indicator Report
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Key_Indicator_Report()
	Select_Link(RACK_GetData("AP_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("AP_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("AP_Data", "Request_Name")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Period").Enter RACK_GetData("AP_Data", "Period")
	If not RACK_GetData("AP_Data", "Include_Invoice") = "" Then
		OracleFlexWindow("Parameters").OracleTextField("Include Invoice Detail").Enter RACK_GetData("AP_Data", "Include_Invoice")
	End If
	If not RACK_GetData("AP_Data", "Max_Vendor_Count") = "" Then
		OracleFlexWindow("Parameters").OracleTextField("Max Vendor Count").Enter RACK_GetData("AP_Data", "Max_Vendor_Count")
	End If
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	OracleFlexWindow("Parameters").Approve
	OracleFormWindow("Submit Request").OracleButton("Submit").Click
    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	RACK_ReportEvent "Validation Screenshot", "Rackspace Key Indicator Report Parameters successfully Enter   ","Screenshot"
	'OracleFormWindow("Submit Request").OracleButton("Submit").Click
	Handle_Oracle_Notification_Forms("OK")
	wait(5)
	Notification_Message = Get_Oracle_Notification_Form_Message()
	Arr=split(Notification_Message, " ")
	Content = Arr(6)
	Print_Request_ID_Arr = Split(Content,")")
	Print_Request_ID = Print_Request_ID_Arr(0)
	Update_Notepad "Print_Request_ID", Print_Request_ID
	Handle_Oracle_Notification_Forms("No")
	Verify_Request_Status(Print_Request_ID)
End Function
'#####################################################################################################################
'Function Description   : Function to run the request Rackspace Uninvoiced Receipt Report
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Uninvoiced_Rec_Report()
	Select_Link(RACK_GetData("AP_Data", "Responsibility_Link"))
	wait(10)
	Select_Link(RACK_GetData("AP_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("AP_Data", "Request_Name")
	If not  RACK_GetData("AP_Data", "SOB_Curr") = ""Then
	OracleFlexWindow("Parameters").OracleTextField("Set of Books Currency").Enter RACK_GetData("AP_Data", "SOB_Curr")
	End If
	If not RACK_GetData("AP_Data", "Period") = "" Then
	OracleFlexWindow("Parameters").OracleTextField("Period Name").Enter RACK_GetData("AP_Data", "Period")
	End If
	If not  RACK_GetData("AP_Data", "Include_Online_Accrual") = ""Then
	OracleFlexWindow("Parameters").OracleTextField("Include Online Accruals").Enter RACK_GetData("AP_Data", "Include_Online_Accrual")
	End If
	If not  RACK_GetData("AP_Data", "Accrued_Receipts") = ""Then
	OracleFlexWindow("Parameters").OracleTextField("Accrued Receipts").Enter RACK_GetData("AP_Data", "Accrued_Receipts")
    End If
	OracleFlexWindow("Parameters").Approve
	OracleFormWindow("Submit Request").OracleButton("Submit").Click
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	Handle_Oracle_Notification_Forms("OK")
	wait(5)
	Notification_Message = Get_Oracle_Notification_Form_Message()
	Arr=split(Notification_Message, " ")
	Content = Arr(6)
	Print_Request_ID_Arr = Split(Content,")")
	Print_Request_ID = Print_Request_ID_Arr(0)
	Update_Notepad "Print_Request_ID", Print_Request_ID
	Handle_Oracle_Notification_Forms("No")
	Verify_Request_Status(Print_Request_ID)
End Function
'#####################################################################################################################
'Function Description   : Function to run Rackspace Invoice Register
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Invoice_Register_Report()
	Select_Link(RACK_GetData("AP_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("AP_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("AP_Data", "Request_Name")
    wait(2)
   OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Parameters").SetFocus
   OracleFlexWindow("Parameters").OracleTextField("From Entered Date").Enter RACK_GetData("AP_Data", "Start_Date")
     wait(2)
   OracleFlexWindow("Parameters").OracleTextField("To Entered Date").Enter RACK_GetData("AP_Data", "End_Date")
     wait(2)
   OracleFlexWindow("Parameters").OracleTextField("Accounting Period").Enter RACK_GetData("AP_Data", "Period")
     wait(2)
   OracleFlexWindow("Parameters").Approve
     wait(2)
	OracleFormWindow("Submit Request").OracleButton("Submit").Click
    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	Handle_Oracle_Notification_Forms("OK")
	wait(5)
	Notification_Message = Get_Oracle_Notification_Form_Message()
	Arr=split(Notification_Message, " ")
	Content = Arr(6)
	Print_Request_ID_Arr = Split(Content,")")
	Print_Request_ID = Print_Request_ID_Arr(0)
	Update_Notepad "Print_Request_ID", Print_Request_ID
	Handle_Oracle_Notification_Forms("No")
	Verify_Request_Status(Print_Request_ID)
End Function
'#####################################################################################################################
'Function Description   : Function to RS US PYBL Processes revised request set
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function RS_US_PYBL_Processes()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("AP_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("AP_Data", "Functionality_Link"))
    If OracleFormWindow("Submit a New Request").Exist(90) Then
		OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request Set").Exist(10) Then
			OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("AP_Data", "Request_Name")
			wait(2)
            OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
            OracleFlexWindow("Parameters").OracleTextField("Option").Enter RACK_GetData("AP_Data", "Option")
             OracleFlexWindow("Parameters").Approve
			 wait(2)
			 OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 2,"Parameters"
             OracleFlexWindow("Parameters").OracleTextField("Ledger").Enter RACK_GetData("AP_Data", "Ledger")
             OracleFlexWindow("Parameters").OracleTextField("End Date").Enter RACK_GetData("AP_Data", "Eff_End_Date")
             OracleFlexWindow("Parameters").OracleTextField("Report").Enter RACK_GetData("AP_Data", "Summarize_Report")
             OracleFlexWindow("Parameters").Approve
			 wait(2)
            OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 3,"Parameters"
            OracleFlexWindow("Parameters").OracleTextField("Ledger").Enter RACK_GetData("AP_Data", "Ledger")
            OracleFlexWindow("Parameters").OracleTextField("End Date").Enter RACK_GetData("AP_Data", "Eff_End_Date")
            OracleFlexWindow("Parameters").Approve
			wait(2)
			OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 4,"Parameters"
            OracleFlexWindow("Parameters").OracleTextField("GL Date").Enter RACK_GetData("AP_Data", "GL_Date")
            OracleFlexWindow("Parameters").OracleTextField("Book").Enter RACK_GetData("AP_Data", "Asset_Book")
            OracleFlexWindow("Parameters").Approve
            wait(2)
			OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 5,"Parameters"
            OracleFlexWindow("Parameters").OracleTextField("Book").Enter RACK_GetData("AP_Data", "Asset_Book")
            OracleFlexWindow("Parameters").Approve
			wait(2)
			OracleFormWindow("Submit Request Set").OracleButton("Submit").Click
            RACK_ReportEvent "Validation Screenshot", "RS US PYBL Processes revised request set Parameters successfully Enter   ","Screenshot"
			Handle_Oracle_Notification_Forms("OK")
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
'#####################################################################################################################
'Function Description   : Function to create payments through Payment Manager
'Input Parameters 	: None
'Return Value    	: None
'Developed : Sravanthi
'##################################################################################################################### 

Function Payment_Batches_Run_Check_Process()
	Select_Link(RACK_GetData("AP_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("AP_Data", "Functionality_Link"))
	wait(10)
	Browser("Oracle Applications Home Page").Page("Payments Dashboard").Link("Payment Process Requests").Click
	wait(2)
	Browser("Oracle Applications Home Page").Page("Payment Process Requests").WebButton("Submit Single Request").Click
	wait(2)
	Browser("Oracle Applications Home Page").Page("Submit Payment Process").WebEdit("CheckrunName").Set RACK_GetData("AP_Data", "Request_Name")
	wait(2)
	Browser("Oracle Applications Home Page").Page("Submit Payment Process").WebEdit("PayFromDate").Set RACK_GetData("AP_Data", "Pay_From_Date")
	wait(2)
	Browser("Oracle Applications Home Page").Page("Submit Payment Process").WebEdit("Payee").Set RACK_GetData("AP_Data", "Trading_Partner")
	wait(2)
	'Browser("Oracle Applications Home Page").Page("Submit Payment Process").Image("Search for Payee").Click
	'wait(12)
	'wait(15)
	'Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame").Image("Quick Select").Click
	wait(5)
	Browser("Oracle Applications Home Page").Page("Submit Payment Process").WebEdit("PaymentMethodName").Set RACK_GetData("AP_Data", "Payment_Type")
	wait(5)
	'Browser("Oracle Applications Home Page").Page("Submit Payment Process").WebEdit("PaymentMethodName").Set RACK_GetData("AP_Data", "Payment_Type")
    Browser("Oracle Applications Home Page").Page("Submit Payment Process").WebEdit("InvoiceBatchName").Set RACK_GetData("AP_Data", "BatchName")
	'Browser("Oracle Applications Home Page").Page("Submit Payment Process").Image("Search for Payment Method").Click
	'wait(15)
	'Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame").Image("Quick Select").Click
    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(5)
	Browser("Oracle Applications Home Page").Page("Submit Payment Process").Link("Payment Attributes").Click
	'wait(15)
	'Browser("Oracle Applications Home Page").Page("Submit Payment Process").WebEdit("UserRateType").Set RACK_GetData("AP_Data", "Rate_Type")
	'Browser("Oracle Applications Home Page").Page("Submit Payment Process").Image("Search for Payment Exchange").Click
	'wait(15)
	'Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame").Image("Quick Select").Click
	wait(5)
	Browser("Oracle Applications Home Page").Page("Submit Payment Process_2").WebEdit("BankAccountName").Set RACK_GetData("AP_Data", "Payment_Bank_Account")
	'Browser("Oracle Applications Home Page").Page("Submit Payment Process_2").Image("Search for Disbursement").Click
	'wait(15)
	'Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame").Image("Quick Select").Click
	wait(5)
	Browser("Oracle Applications Home Page").Page("Submit Payment Process").WebEdit("PaymentProfileName").Set RACK_GetData("AP_Data", "Payment_Process_Profile")
	'Browser("Oracle Applications Home Page").Page("Submit Payment Process").Image("Search for Payment Process").Click
	wait(5)
	Browser("Oracle Applications Home Page").Page("Submit Payment Process").WebEdit("UserRateType").Set RACK_GetData("AP_Data", "Rate_Type")
	wait(10)
    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
'	wait(2)
'    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	Browser("Oracle Applications Home Page").Page("Submit Payment Process").WebButton("Submit").Click

'	Confirmation_Message = Browser("Oracle Applications Home Page").Page("Payment Process Requests_2").WebTable("Confirmation").GetCellData(1,1)
    wait(5)
	Browser("Oracle Applications Home Page").Page("Payment Process Requests").WebEdit("SearchCheckrunName").Set RACK_GetData("AP_Data", "Request_Name")
	'Browser("Oracle Applications Home Page").Page("Payment Process Requests").Image("Search for Payment Process").Click
	'wait(5)
	'Browser("Oracle Applications R12").Page("Search and Select List").Frame("Frame").Image("Quick Select").Click
	Browser("Oracle Applications Home Page").Page("Payment Process Requests").WebButton("Go").Click
	wait(5)
	RACK_ReportEvent "Validation Screenshot", "Create Payments is created sucessfully ","Pass"
End Function

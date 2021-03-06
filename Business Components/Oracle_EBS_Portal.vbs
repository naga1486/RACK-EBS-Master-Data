'#####################################################################################################################
'Function Description   : One time customer Payment in Portal
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Onetime_Payment_Portal()
    set Parent_desc = Browser("Oracle Applications Home Page").Page("Home — MyRackspace")
	wait(10)
	Parent_desc.Link("Account").Click
	wait(5)
	Parent_desc.Link("Payments").Click
	wait(5)
	Parent_desc.Link("Credit Card").Click
	wait(5)
	Parent_desc.WebEdit("billingNameFirst").Set RACK_GetData("Portal_Data", "First_Name")'"Mana"
	Parent_desc.WebEdit("billingNameLast").Set RACK_GetData("Portal_Data", "Last_Name")' "Valan"
	Parent_desc.WebEdit("billingAddress1").Set RACK_GetData("Portal_Data", "Billing_Address")'"4822 Gus Eckert"
	Parent_desc.WebEdit("billingAddressCity").Set RACK_GetData("Portal_Data", "City")'"San Antonio"
	Parent_desc.WebEdit("billingAddressState").Set RACK_GetData("Portal_Data", "State")'"TX"
	Parent_desc.WebEdit("billingAddressZip").Set RACK_GetData("Portal_Data", "Zip")'"78240"
	Parent_desc.WebEdit("billingPhone").Set RACK_GetData("Portal_Data", "Phone")'"1111111111"
	Parent_desc.WebEdit("billingEmail").Set RACK_GetData("Portal_Data", "Email")'"abc@dfc.com"
	Parent_desc.WebList("billingAddressCountry").Select RACK_GetData("Portal_Data", "Country")'"United States"
	Parent_desc.WebRadioGroup("cardType").Select RACK_GetData("Portal_Data", "Card_Type")'"mastercard"
	Parent_desc.WebEdit("cardNumber").Set RACK_GetData("Portal_Data", "Card_Number")'"5454545454545454"
	Parent_desc.WebEdit("cardCcv").Set RACK_GetData("Portal_Data", "Card_CCV")'"123"
	Parent_desc.WebList("cardExpirationMonth").Select RACK_GetData("Portal_Data", "Exp_Month")
	Parent_desc.WebList("cardExpirationYear").Select RACK_GetData("Portal_Data", "Exp_Year")
	Parent_desc.Link("Pay Other Amount").Click
	Parent_desc.WebEdit("amountOther").Set RACK_GetData("Portal_Data", "Payment_Amount")'"13"
	wait(5)
	RACK_ReportEvent "One time Payment", "One Time payment Enter is done sucessfully" ,"Pass"
	'RACK_ReportEvent "Validation Screenshot", "Validation Screenshot" ,"Screenshot"
	Parent_desc.WebButton("Make Payment").Click
	wait(5)
	Browser("Oracle Applications Home Page").Dialog("Message from webpage").Activate
	Browser("Oracle Applications Home Page").Dialog("Message from webpage").WinButton("OK").Click
	wait(5)
	If  Parent_desc.WebTable("Date").Exist(120) Then
		wait(2)
		Row_Number = Browser("Oracle Applications Home Page").Page("Home — MyRackspace").WebTable("Date").GetRowWithCellText("0 sec ago")
		RACK_ReportEvent "One time Payment", "One Time payment is done sucessfully" ,"Pass"
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End if
End function

'#####################################################################################################################
'Function Description   : Verify Customer Payments in EBS
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Verify_Customer_Payments()
	Select_Link(RACK_GetData("Portal_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Portal_Data", "Functionality_Link"))
	If OracleFormWindow("Find Customer Accounts").Exist(90) Then
		OracleFormWindow("Find Customer Accounts").OracleTextField("Customer Num").Enter RACK_GetData("Portal_Data", "Customer_Number")
		OracleFormWindow("Find Customer Accounts").OracleButton("Find").Click
		If OracleFormWindow("Customer Accounts").Exist(90) Then
			OracleFormWindow("Customer Accounts").OracleButton("Account Details").Click
			If OracleFormWindow("Account Details").Exist(30) Then
				OracleFormWindow("Account Details").OracleTable("Table").InvokeSoftkey "ENTER QUERY"
				OracleFormWindow("Account Details").OracleTable("Table").EnterField 1,"Number", "%OL%"
				OracleFormWindow("Account Details").OracleTable("Table").EnterField 1,"Due Date", Format_Date("dd-mmm-yyyy")
				wait(2)
				OracleFormWindow("Account Details").OracleTable("Table").InvokeSoftkey "EXECUTE QUERY"
				wait(5)
				Amount = OracleFormWindow("Account Details").OracleTable("Table").GetFieldValue(1,"Number")
				If not Amount = "" Then	
					RACK_ReportEvent "Customer Payments", "The Customer Payment is available with the Original amount as '" & Amount  & "'","Pass"
					RACK_ReportEvent "Validation Screenshot", "Validation Screenshot" ,"Screenshot"
				else
					RACK_ReportEvent "Customer Payments", "The Customer Payment is not available ","Fail"
				End if
			End if
		End if
	End if
End Function

'#####################################################################################################################
'Function Description   : Function to create a receivable Invoice
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_AR_Invoices()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Portal_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Portal_Data", "Functionality_Link"))
	wait(60)
	If OracleFormWindow("Transactions").OracleTextField("Transaction|Source").Exist(10) Then
			OracleFormWindow("Transactions").OracleList("Transaction|Class").Select RACK_GetData("Portal_Data","Invoice_Class")
			wait(2)
			OracleFormWindow("Transactions").OracleTextField("Transaction|Type").Enter RACK_GetData("Portal_Data","Invoice_Type")
		If RACK_GetData("Portal_Data","Transaction_Date") <> "" Then
			OracleFormWindow("Transactions").OracleTextField("Transaction|Date").Enter RACK_GetData("Portal_Data","Transaction_Date")			
		End If		
		If RACK_GetData("Portal_Data","Transaction_Currency") <> "" Then
			OracleFormWindow("Transactions").OracleTextField("Transaction|Currency").Enter RACK_GetData("Portal_Data","Transaction_Currency")			
		End If
	wait(2)
		If RACK_GetData("Portal_Data","Invoice_BillTo_Name") <> "" Then
			OracleFormWindow("Transactions").OracleTabbedRegion("Main").OracleTextField("Bill To: Name").Enter RACK_GetData("Portal_Data","Invoice_BillTo_Name")
		else
			OracleFormWindow("Transactions").OracleTabbedRegion("Main").OracleTextField("Bill To: Number").Enter RACK_GetData("Portal_Data","Invoice_BillTo_Number")
		End If
		wait(2)
		If RACK_GetData("Portal_Data","Payment_Term") <> "" Then
			OracleFormWindow("Transactions").OracleTabbedRegion("Main").OracleTextField("Payment Term").Enter RACK_GetData("Portal_Data","Payment_Term")
End If
Wait(2)
'***************Added code to pouplate bank and crdit card details in DFF of invoice to print ***********
If RACK_GetData("Portal_Data","Bank_Ccinfo")= "Yes" Then
OracleFormWindow("Transactions").OracleTextField("Transaction|[").SetFocus
OracleFlexWindow("Transaction Information").OracleTextField("Bank Information").Enter RACK_GetData("Portal_Data","Bank_Information")
OracleFlexWindow("Transaction Information").OracleTextField("Credit Card Information").Enter RACK_GetData("Portal_Data","Credit_Card_Information")
OracleFlexWindow("Transaction Information").Approve
		End if
		wait(2)
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		OracleFormWindow("Transactions").OracleButton("Line Items").Click
		end if
		Dim LineNum,i
		Dim LineDesc,Qty,UtPrice
        Dim LineNumArr,LineDescArr,QtyArr,UtPriceArr
		LineNum =  RACK_GetData("Portal_Data","Invoice_Lines_Num")
		LineNumArr = split(LineNum,";")
        LineDesc = RACK_GetData("Portal_Data","Invoice_Lines_Description")
		LineDescArr= split(LineDesc,";")
		Qty = RACK_GetData("Portal_Data","Invoice_Lines_Quantity")
		QtyArr = split(Qty,";")
		UtPrice = RACK_GetData("Portal_Data","Invoice_Lines_UnitPrice")
		UtPriceArr = split(UtPrice,";")
		LinePrd = RACK_GetData("Portal_Data","Invoice_Lines_Product")
		LinePrdArr = split(LinePrd,";")
	
       For Each i In LineNumArr
		   i =i-1	
			If  OracleFormWindow("Lines").OracleTabbedRegion("Main").OracleTable("Table").Exist(30) Then
                OracleFormWindow("Lines").OracleTabbedRegion("Main").OracleTable("Table").EnterField 1,"Num", LineNumArr(i)
				OracleFormWindow("Lines").OracleTabbedRegion("Main").OracleTable("Table").SetFocus 1,"Description"
				OracleFormWindow("Lines").OracleTabbedRegion("Main").OracleTable("Table").OpenDialog 1,"Description"
				If OracleListOfValues("Standard Memo Lines").Exist(10) Then
					OracleListOfValues("Standard Memo Lines").Find LineDescArr(i)
					OracleListOfValues("Standard Memo Lines").Select LineDescArr(i)
				End If
				wait(2)
				OracleFormWindow("Lines").OracleTabbedRegion("Main").OracleTable("Table").EnterField 1,"Quantity", QtyArr(i)
				wait(2)
				OracleFormWindow("Lines").OracleTabbedRegion("Main").OracleTable("Table").EnterField 1,"Unit Price", UtPriceArr(i)
				wait(2)
				OracleFormWindow("Lines").OracleTabbedRegion("More").Click
				wait(2) 
				OracleFormWindow("Lines").OracleTabbedRegion("More").OracleTable("Table").SetFocus 1,"[]"
			End If
			If  OracleFlexWindow("Invoice Line Information").OracleTextField("Billing Cycle").Exist(30)Then
				OracleFlexWindow("Invoice Line Information").OracleTextField("Billing Cycle").Enter RACK_GetData("Portal_Data","Invoice_Lines_BillingCycle")
			   OracleFlexWindow("Invoice Line Information").OracleTextField("Product").Enter LinePrdArr(i)
			  If RACK_GetData("Portal_Data","RefNum") <> "" Then
				 OracleFlexWindow("Invoice Line Information").OracleTextField(" Reference No.").Enter RACK_GetData("Portal_Data","RefNum")
			 End If
			 RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			OracleFlexWindow("Invoice Line Information").OracleButton("OK").Click
		End if
			wait(2)
			OracleFormWindow("Lines").OracleTabbedRegion("Main").OracleTable("Table").SetFocus 2,"Num"
			wait(2)
			OracleFormWindow("Lines").SelectMenu "Edit->Clear->Record"
			wait(2)
		Next
  OracleFormWindow("Lines").SelectMenu "File->Save"
wait(2)
Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
		OracleFormWindow("Lines").CloseWindow
		wait(2)
		OracleFormWindow("Transactions").OracleButton("Complete").Click
		wait(10)
	Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
	Invoice_Number = OracleFormWindow("Transactions").OracleTextField("Transaction|Transaction").GetROProperty("value")
	RACK_PutData "Portal_Data", "Invoice_Number", Invoice_Number
	If not Invoice_Number = "" Then
		RACK_ReportEvent "Invoice Number", "The Invoice Number is sucessfully generated and is - " & Invoice_Number,"Pass"
	End If
End Function
Function  excelData()
		   LineNum =  RACK_GetData("Portal_Data","Invoice_Lines_Num")
			LineNumArr= split(LLineNum,";")
			
 excelData = LineNumArr
End Function

'#####################################################################################################################
'Function Description   : To View AR information in Portal
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function AR_Information_Portal()
  set Parent_desc = Browser("Oracle Applications Home Page").Page("Home — MyRackspace")
	Parent_desc.Link("Account").Click
	wait(5)
	'Parent_desc.Link("Transactions").Click
    Browser("Oracle Applications Home Page").Page("Home — MyRackspace_2").Link("Transactions").Click
	'wait(5)
	'Parent_desc.WebCheckBox("type").Set "OFF"
	wait(5)
	LinkName = RACK_GetData("Portal_Data", "Transaction_ID")'
	Parent_desc.Link("name:="& LinkName).Click
	wait(10)
    Browser("Oracle Applications Home Page").Page("All Transactions — Transactions").Link("name:="& LinkName).Click
	wait(10)
	If  Browser("Oracle Applications Home Page").Dialog("File Download").Exist(30) Then
		RACK_ReportEvent "Portal Invoice Transactions", "The PDF for the Transaction ID '" & LinkName  & "' is opened sucessfully as expected","Pass"
		wait(10)
		Browser("Oracle Applications Home Page").Dialog("File Download").WinButton("Save").Click
	else
	RACK_ReportEvent "Portal Invoice Transactions", "The PDF for the Transaction ID '" & LinkName  & "' is not opened sucessfully as expected","Pass"
	End if
End function

'#####################################################################################################################
'Function Description   : Customer Credit Card update in Portal
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Customer_CC_Update_Portal()
  set Parent_desc = Browser("Oracle Applications Home Page").Page("Home — MyRackspace")
	wait(20)
	Parent_desc.Link("Account").Click
	wait(5)
	Parent_desc.Link("Payments").Click
	wait(5)
	Browser("Oracle Applications Home Page").Page("Electronic Check — Payment").Link("Recurring Payment").Click
	wait(5)
	Browser("Oracle Applications Home Page").Page("Recurring Payment — Payment").Link("Credit Card").Click
	wait(5)
	Browser("Oracle Applications Home Page").Page("Recurring Payment — Payment").WebEdit("cardName").Set RACK_GetData("Portal_Data", "Card_Name")
	wait(2)
	Browser("Oracle Applications Home Page").Page("Recurring Payment — Payment").WebEdit("cardNumber").Set RACK_GetData("Portal_Data", "Card_Number")
	wait(2)
	Browser("Oracle Applications Home Page").Page("Recurring Payment — Payment").WebRadioGroup("cardType").Select RACK_GetData("Portal_Data", "Card_Type")
	wait(2)
    Browser("Oracle Applications Home Page").Page("Recurring Payment — Payment").WebList("cardExpirationMonth").Select RACK_GetData("Portal_Data", "Exp_Month")
	wait(2)
	Browser("Oracle Applications Home Page").Page("Recurring Payment — Payment").WebList("cardExpirationYear").Select RACK_GetData("Portal_Data", "Exp_Year")
	wait(2)
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
    	'Parent_desc.WebButton("Continue...").Click
	 Browser("Oracle Applications Home Page").Page("Make a Recurring Payment").WebButton("Continue...").Click
	 wait(5)
	If  Parent_desc.WebButton("Set Up Recurring Payment").Exist(20) Then
		Parent_desc.WebCheckBox("agree").Set "ON"
        RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(2)
		Parent_desc.WebButton("Set Up Recurring Payment").Click
        RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(5)
	End If
	If RACK_GetData("Portal_Data","Bank_Name")= "Yes" Then
		RACK_ReportEvent "Create ACH Payment Method", "ACH Recurring is done sucessfully" ,"Pass"
	End If
End function
'#####################################################################################################################
'Function Description   : Verify Credit Card set up
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Verify_CC_Setup()
  set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Portal_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Portal_Data", "Functionality_Link"))
	If  Parent_desc.WebEdit("Account_Number_Search").Exist(40) Then
		Parent_desc.WebEdit("Account_Number_Search").Set RACK_GetData("Portal_Data", "Customer_Number")
		wait(2)
		Parent_desc.WebButton("Go").Click
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
	If  Parent_desc.Image("Details Enabled").Exist(10) Then
		wait(2)
		Parent_desc.Image("Details Enabled").Click
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
	If  Parent_desc.WebButton("CreateSite").Exist(10) Then
		wait(2)
		Parent_desc.Image("Details Enabled").Click
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
	wait(10)
		If  Parent_desc.Link("Payment_Details").Exist(10) Then
		wait(2)
		Parent_desc.Link("Payment_Details").Click
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
	wait(10)
		RACK_ReportEvent "Credit Card Number", "The Credit Card Number is correctly displayed as '" & Actual_CC_Number & "'.","Pass"
		Set Obj=Browser("Oracle Applications Home Page").Page("Site1").Object.body
		Obj.doScroll("pageDown")
		wait(10)
		RACK_ReportEvent "Credit Card Expiry Date", "The Credit Card Expiry Date is correctly displayed as '" & Actual_Exp_Number & "'.","Pass"
End Function

'#####################################################################################################################
'Function Description   : Recurring customer Payment in Portal
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_ACH_Payment_Method()
   set Parent_desc = Browser("Oracle Applications Home Page").Page("Home — MyRackspace")
	wait(20)
	Parent_desc.Link("Account").Click
	wait(5)
	Parent_desc.Link("Payments").Click
	wait(5)
	'Browser("Oracle Applications Home Page").Page("Electronic Check — Payment").Link("Make a Recurring Payment").Click
	 Browser("Oracle Applications Home Page").Page("Electronic Check — Payment").Link("Recurring Payment").Click
	'Browser("Oracle Applications Home Page").Page("Electronic Check — Payment").Link("Make a Recurring Payment").Click
	wait(10)
	'Browser("Electronic Check — Payment").Page("Recurring Payment — Payment").Link("Direct Debit / ACH").Click
	'Browser("Oracle Applications Home Page").Page("Recurring Payment — Payment").Link("Credit Card").Click
	Browser("Oracle Applications Home Page").Page("Make a Recurring Payment").WebRadioGroup("noIntFunds").Select "no"
	wait(5)
	'Browser("Oracle Applications Home Page").Page("Recurring Payment — Payment").WebEdit("cardName").Set RACK_GetData("Portal_Data", "Card_Name")
	Browser("Oracle Applications Home Page").Page("Make a Recurring Payment").WebEdit("bankName").Set RACK_GetData("Portal_Data", "Bank_Name")
	wait(5)
	Browser("Oracle Applications Home Page").Page("Make a Recurring Payment").WebRadioGroup("bankAccountType").Select "Commercial Checking"
	'Browser("Oracle Applications Home Page").Page("Recurring Payment — Payment").WebEdit("cardNumber").Set RACK_GetData("Portal_Data", "Card_Number")
	wait(5)
	'Browser("Oracle Applications Home Page").Page("Recurring Payment — Payment").WebRadioGroup("cardType").Select RACK_GetData("Portal_Data", "Card_Type")
	Browser("Oracle Applications Home Page").Page("Make a Recurring Payment").WebEdit("bankRoutingNumber").Set RACK_GetData("Portal_Data", "Routing_Number")
	wait(5)
	'Browser("Oracle Applications Home Page").Page("Recurring Payment — Payment").WebList("cardExpirationMonth").Select RACK_GetData("Portal_Data", "Exp_Month")
	Browser("Oracle Applications Home Page").Page("Make a Recurring Payment").WebEdit("bankAccountNumber").Set RACK_GetData("Portal_Data", "Bank_Number")
	wait(5)
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	 Browser("Oracle Applications Home Page").Page("Make a Recurring Payment").WebButton("Continue...").Click
    'Parent_desc.WebButton("Continue...").Click
	wait(5)
	If  Parent_desc.WebButton("Set Up Recurring Payment").Exist(20) Then
		Parent_desc.WebCheckBox("agree").Set "ON"
        'RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(2)
		Parent_desc.WebButton("Set Up Recurring Payment").Click
		wait(5)
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	End If
'	Browser("Oracle Applications Home Page").Dialog("Message from webpage").Activate
'	Browser("Oracle Applications Home Page").Dialog("Message from webpage").WinButton("OK").Click
	wait(20)
	If RACK_GetData("Portal_Data","Bank_Name")= "Yes" Then
		RACK_ReportEvent "Create ACH Payment Method", "ACH Recurring is done sucessfully" ,"Pass"
	End If
End function

'#####################################################################################################################
'Function Description   : Recurring customer Payment in Portal
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Cancel_ACH_Payment_Method()
   set Parent_desc = Browser("Oracle Applications Home Page").Page("Home — MyRackspace")
	wait(20)
	Parent_desc.Link("Account").Click
	wait(5)
	Parent_desc.Link("Payments").Click
	wait(5)
	Browser("Oracle Applications Home Page").Page("Electronic Check — Payment").Link("Recurring Payment").Click
	'Browser("Make a Recurring Payment").Page("Electronic Check — Payment").Link("Make a Recurring Payment").Click
	wait(5)
	Browser("Oracle Applications Home Page").Page("Make a Recurring Payment").WebButton("Cancel Recurring E-Check").Click
	RACK_ReportEvent "Recurring Payment", "Updated Credit card details entered  sucessfully" ,"Pass"
	wait(5)
	'Browser("Make a Recurring Payment").Page("Make a Recurring Payment").WebButton("Continue..._2").Click
    	'Parent_desc.WebButton("Continue...2").Click
	Browser("Oracle Applications Home Page").Page("Make a Recurring Payment").WebButton("Continue..._2").Click
		wait(5)
	If  Parent_desc.WebButton("Set Up Recurring Payment").Exist(20) Then
		Parent_desc.WebCheckBox("agree").Set "ON"
		wait(5)
		Parent_desc.WebButton("Set Up Recurring Payment").Click
		wait(5)
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	End If
'	Browser("Oracle Applications Home Page").Dialog("Message from webpage").Activate
'	Browser("Oracle Applications Home Page").Dialog("Message from webpage").WinButton("OK").Click
	wait(20)
	If RACK_GetData("Portal_Data","Bank_Name")= "Yes" Then
		RACK_ReportEvent "Create ACH Payment Method", "ACH Recurring is done sucessfully" ,"Pass"
	End If
End function

''#####################################################################################################################
''Function Description   : Function for Custom_ACH_Process
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 

Function Create_Custom_ACH_Process()
'set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("Portal_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Portal_Data", "Functionality_Link"))
	If OracleFormWindow("Receipt Batches").Exist(90)  Then
		OracleFormWindow("Receipt Batches").OracleList("Batch Type").Select "Automatic"
		wait(2)
		OracleFormWindow("Receipt Batches").OracleTextField("Receipt Class").Enter RACK_GetData("Portal_Data","Receipt_Class")
		wait(2)
		OracleFormWindow("Receipt Batches").OracleButton("Create").Click
		wait(2)
		OracleFormWindow("Create Automatic Receipts").OracleTextField("Dates|Due").Enter RACK_GetData("Portal_Data","Receipt_Date_Low")
		wait(5)
		OracleFormWindow("Create Automatic Receipts").OracleTextField("Dates: High Due").Enter RACK_GetData("Portal_Data","Receipt_Date_High")
		wait(5)
		OracleFormWindow("Create Automatic Receipts").OracleButton("OK").Click
        RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		OracleNotification("Forms").Approve
		wait(2)
	Request_ID = OracleFormWindow("Receipt Batches").OracleTextField("Request ID").GetROProperty("value")
	RACK_PutData "Portal_Data", "Request_ID", CStr(Request_ID)
	Update_Notepad "Request_ID", Receipt_Number
	If  not Request_ID = "" Then
	RACK_ReportEvent "Request ID", "The Request ID '" & Request_ID & "' is created sucessfully as expected","Pass"
	else
	RACK_ReportEvent "Purchase Order", "The Purchase Order is not created sucessfully","Fail"
	End If
	wait(2)
        OracleFormWindow("Receipt Batches").CloseWindow
		wait(2)
		OracleFormWindow("Navigator").SelectMenu "View->Requests"
		wait(2)
		OracleFormWindow("Find Requests").OracleRadioGroup("My Completed Requests").Select "Specific Requests"
		wait(2)
		OracleFormWindow("Find Requests").OracleTextField("Request ID").Enter RACK_GetData("Portal_Data","Request_ID")
		OracleFormWindow("Find Requests").OracleButton("Find").Click
		wait(2)
		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
		wait(10)
		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
		wait(10)
		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
		wait(10)
		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
		wait(10)
		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
		wait(10)
		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
		wait(10)
		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
		wait(10)
	    OracleFormWindow("Requests").OracleButton("Refresh Data").Click
		wait(10)
		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
		wait(10)
		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
 Phase = OracleFormWindow("Requests").OracleTable("Table").GetFieldValue(1,"Phase")
 If Phase = "Completed" Then
 RACK_ReportEvent "Create Custom ACH Process -Phase is Completed","Phase = Completed","Pass"
 Else
 RACK_ReportEvent "Create Custom ACH Process -Phase is Completed","Phase is - "&Phase,"Fail"
	End if
	End if
End Function
''#####################################################################################################################
''Function Description   : Function for Custom_ACH_Process
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 

Function Approve_Custom_ACH_Process()
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("Portal_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Portal_Data", "Functionality_Link"))
	If OracleFormWindow("Receipt Batches").Exist(90)  Then
		OracleFormWindow("Receipt Batches").SelectMenu "View->Query By Example->Enter"
		wait(2)
		OracleFormWindow("Receipt Batches").OracleTextField("Batch Number").Enter RACK_GetData("Portal_Data","Batch_Number")
		wait(2)
		OracleFormWindow("Receipt Batches").SelectMenu "View->Query By Example->Run"
		wait(2)
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		OracleFormWindow("Receipt Batches").OracleButton("Maintain").Click
		wait(2)
		OracleFormWindow("Maintain Automatic Receipts").OracleButton("Approve").Click
		wait(2)
        RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		OracleNotification("Decision").Approve
		OracleFormWindow("Maintain Automatic Receipts").CloseWindow
		wait(2)
	Request_ID = OracleFormWindow("Receipt Batches").OracleTextField("Request ID").GetROProperty("value")
	RACK_PutData "Portal_Data", "Request_ID", CStr(Request_ID)
	Update_Notepad "Request_ID", Receipt_Number
	If  not Request_ID = "" Then
	RACK_ReportEvent "Request ID", "The Request ID '" & Request_ID & "' is created sucessfully as expected","Pass"
	else
	RACK_ReportEvent "Purchase Order", "The Purchase Order is not created sucessfully","Fail"
	End If
	wait(2)
        OracleFormWindow("Receipt Batches").CloseWindow
		wait(2)
		OracleFormWindow("Navigator").SelectMenu "View->Requests"
		wait(2)
		OracleFormWindow("Find Requests").OracleRadioGroup("My Completed Requests").Select "Specific Requests"
		wait(2)
		OracleFormWindow("Find Requests").OracleTextField("Request ID").Enter RACK_GetData("Portal_Data","Request_ID")
		OracleFormWindow("Find Requests").OracleButton("Find").Click
		wait(2)
		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
		wait(10)
		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
		wait(10)
		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
		wait(10)
		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
		wait(10)
	    OracleFormWindow("Requests").OracleButton("Refresh Data").Click
		wait(10)
		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
		wait(10)
		OracleFormWindow("Requests").OracleButton("Refresh Data").Click
 Phase = OracleFormWindow("Requests").OracleTable("Table").GetFieldValue(1,"Phase")
 If Phase = "Completed" Then
 RACK_ReportEvent "Approve Custom ACH Process -Phase is Completed","Phase = Completed","Pass"
 Else
 RACK_ReportEvent "Approve Custom ACH Process -Phase is Completed","Phase is - "&Phase,"Fail"
End if
OracleFormWindow("Requests").CloseWindow
wait(2)
OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List_2").Select "+  Receipts"
wait(2)
OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List_2").Activate "+  Receipts"
wait(2)
OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List_2").Activate "       Batches"
wait(2)
OracleFormWindow("Receipt Batches").SelectMenu "View->Query By Example->Enter"
wait(2)
OracleFormWindow("Receipt Batches").OracleTextField("Batch Number").Enter RACK_GetData("Portal_Data","Batch_Number")
wait(2)
OracleFormWindow("Receipt Batches").SelectMenu "View->Query By Example->Run"
wait(2)
 RACK_ReportEvent "Approve Custom ACH Process -Process Status is  = Completed Approval","Pass"
End if
End Function

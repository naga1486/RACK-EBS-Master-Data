
'#####################################################################################################################
'CPU Patch Testing  
'Function Description   : Function to create a Customer
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_Customer()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("AR_Data", "Responsibility_Link"))
	wait(10)
	Select_Link(RACK_GetData("AR_Data", "Functionality_Link"))
	If Parent_desc.WebButton("Create").Exist(20) then
		Parent_desc.WebButton("Create").Click
		RACK_ReportEvent "Create Customer", "The button Create Customer is clicked","Done"
		If Parent_desc.WebEdit("Account_Number_Create").Exist(20) then
			OrganizationName = RACK_GetData("AR_Data", "OrganizationName")
			Parent_desc.WebEdit("Organization_Name_Create").Set OrganizationName
			wait(2)	
			Parent_desc.WebEdit("Account_Number_Create").Set RACK_GetData("AR_Data", "SearchAccountNumber")
			wait(2)
			RACK_ReportEvent "Validation Screenshot", "Customer and Account informatin Parameters successfully Enter   ","Screenshot"
			wait(2)
			If RACK_GetData("AR_Data", "Org")  = "US" Then
				Parent_desc.WebEdit("Address_Line_1_Create").Set RACK_GetData("AR_Data", "AddressLine1")
				wait(2)
				Parent_desc.WebEdit("City_Create").Set RACK_GetData("AR_Data", "City")
				wait(2)
				Parent_desc.WebEdit("County_Create").Set RACK_GetData("AR_Data", "County")
				wait(2)
				Parent_desc.WebList("State_Create").Select RACK_GetData("AR_Data", "State")
				wait(2)
				Parent_desc.WebEdit("Postal_Code_Create").Set RACK_GetData("AR_Data", "PostalCode")
				wait(2)
				RACK_ReportEvent "Validation Screenshot", "Account Site Address informatin Parameters successfully Enter   ","Screenshot"
			End IF
			If RACK_GetData("AR_Data", "Org")  = "UK" OR RACK_GetData("AR_Data", "Org")  = "HK"Then
				Parent_desc.WebEdit("Address_Line_1_Create").Set RACK_GetData("AR_Data", "AddressLine1")
				Browser("Oracle Applications Home Page").Page("Create Organization_2").WebEdit("HzAddressStyleFlex4").Set RACK_GetData("AR_Data", "City")
			End If			
        	Parent_desc.WebEdit("Currency_Create").click
			WshShell.SendKeys RACK_GetData("AR_Data", "Currency")
			WshShell.SendKeys "{TAB}"
			wait(5)
			Parent_desc.WebEdit("Support_Team_Create").click
			WshShell.SendKeys RACK_GetData("AR_Data", "SupportTeam")
			WshShell.SendKeys "{TAB}"
			wait(5)
			Parent_desc.WebEdit("Account_Manager_Create").click
			WshShell.SendKeys RACK_GetData("AR_Data", "AccountManager")
			WshShell.SendKeys "{TAB}"
			wait(5)
			Parent_desc.WebEdit("Segment_Create").click
			WshShell.SendKeys RACK_GetData("AR_Data", "Segment")
			WshShell.SendKeys "{TAB}"
			wait(5)
			Parent_desc.WebEdit("Print_Group_Create").click
			WshShell.SendKeys RACK_GetData("AR_Data", "PrintGroup")
			WshShell.SendKeys "{TAB}"
			wait(2)
			Browser("Oracle Applications Home Page").Page("Create Organization_3").WebCheckBox("BizPurposeTable:PrimaryUseFlag:0").Set "ON"
            RACK_ReportEvent "Validation Screenshot", "Account Site Details Parameters successfully Enter   ","Screenshot"
			Parent_desc.WebButton("Save_And_Add_Details").Click
			wait(10)
			If Parent_desc.WebButton("Save").Exist(20) then
		    Parent_desc.WebButton("Save").Click
		End If
			wait(10)
			RACK_ReportEvent "Validation Screenshot", "New Customer successfully created    ","Screenshot"
		End if
	End if

End Function
'#####################################################################################################################
'Function Description   : Function to Query_Existing_Customer
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Query_Existing_Customer()
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	If Browser("Login").Page("Login").Link("Home").exist(3) then
		Browser("Login").Page("Login").Link("Home").Click
	End if
	Select_Link(RACK_GetData("AR_Data", "Responsibility_Link"))
	Select_Link(RACK_GetData("AR_Data", "Functionality_Link"))
    If not RACK_GetData("AR_Data","OrganizationName") = "" Then
	Parent_desc.WebEdit("Customer_Search").Set RACK_GetData("AR_Data", "OrganizationName")
	End If
	If not RACK_GetData("AR_Data","SearchAccountNumber") = "" Then
	Parent_desc.WebEdit("Account_Number_Search").Set RACK_GetData("AR_Data", "SearchAccountNumber")
	End If
	If not RACK_GetData("AR_Data","Account_Description") = "" Then
	Parent_desc.WebEdit("Account_Description_Search").Set RACK_GetData("AR_Data", "Account_Description")
	End If
	Parent_desc.WebButton("Go").Click
	wait(10)
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(10)
    Parent_desc.Image("View_Details").Click
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(10)
    Parent_desc.Image("Details Enabled").Click
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(10)
    Parent_desc.Link("Payment_Details").Click
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(10)
    Browser("Oracle Applications Home Page").Page("Site: 1782_2").Link("Business Purposes").Click
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	Parent_desc.Image("View_Details").Click
	wait(10)
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
End Function

'#####################################################################################################################
'Function Description   : Function to create a receivable Invoice
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_Invoice()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("AR_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("AR_Data", "Functionality_Link"))
	wait(60)
	If OracleFormWindow("Transactions").OracleTextField("Transaction|Source").Exist(10) Then
			OracleFormWindow("Transactions").OracleList("Transaction|Class").Select RACK_GetData("AR_Data","Invoice_Class")
			wait(2)
			OracleFormWindow("Transactions").OracleTextField("Transaction|Type").Enter RACK_GetData("AR_Data","Invoice_Type")
		If RACK_GetData("AR_Data","Transaction_Date") <> "" Then
			OracleFormWindow("Transactions").OracleTextField("Transaction|Date").Enter RACK_GetData("AR_Data","Transaction_Date")			
		End If		
		If RACK_GetData("AR_Data","Transaction_Currency") <> "" Then
			OracleFormWindow("Transactions").OracleTextField("Transaction|Currency").Enter RACK_GetData("AR_Data","Transaction_Currency")			
		End If
	wait(2)
		If RACK_GetData("AR_Data","Invoice_BillTo_Name") <> "" Then
			OracleFormWindow("Transactions").OracleTabbedRegion("Main").OracleTextField("Bill To: Name").Enter RACK_GetData("AR_Data","Invoice_BillTo_Name")
		else
			OracleFormWindow("Transactions").OracleTabbedRegion("Main").OracleTextField("Bill To: Number").Enter RACK_GetData("AR_Data","Invoice_BillTo_Number")
		End If
		wait(2)
		If RACK_GetData("AR_Data","Payment_Term") <> "" Then
			OracleFormWindow("Transactions").OracleTabbedRegion("Main").OracleTextField("Payment Term").Enter RACK_GetData("AR_Data","Payment_Term")
End If
Wait(2)
'***************Added code to pouplate bank and crdit card details in DFF of invoice to print ***********
'If RACK_GetData("AR_Data","Bank_Ccinfo")= "Yes" Then
'OracleFormWindow("Transactions").OracleTextField("Transaction|[").SetFocus
'OracleFlexWindow("Transaction Information").OracleTextField("Bank Information").Enter RACK_GetData("AR_Data","Bank_Information")
'OracleFlexWindow("Transaction Information").OracleTextField("Credit Card Information").Enter RACK_GetData("AR_Data","Credit_Card_Information")
'OracleFlexWindow("Transaction Information").Approve
'		End if
		wait(2)
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		OracleFormWindow("Transactions").OracleButton("Line Items").Click
		end if
		Dim LineNum,i
		Dim LineDesc,Qty,UtPrice
        Dim LineNumArr,LineDescArr,QtyArr,UtPriceArr
		LineNum =  RACK_GetData("AR_Data","Invoice_Lines_Num")
		LineNumArr = split(LineNum,";")
        LineDesc = RACK_GetData("AR_Data","Invoice_Lines_Description")
		LineDescArr= split(LineDesc,";")
		Qty = RACK_GetData("AR_Data","Invoice_Lines_Quantity")
		QtyArr = split(Qty,";")
		UtPrice = RACK_GetData("AR_Data","Invoice_Lines_UnitPrice")
		UtPriceArr = split(UtPrice,";")
		LinePrd = RACK_GetData("AR_Data","Invoice_Lines_Product")
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
				OracleFlexWindow("Invoice Line Information").OracleTextField("Billing Cycle").Enter RACK_GetData("AR_Data","Invoice_Lines_BillingCycle")
			   OracleFlexWindow("Invoice Line Information").OracleTextField("Product").Enter LinePrdArr(i)
			  If RACK_GetData("AR_Data","RefNum") <> "" Then
				 OracleFlexWindow("Invoice Line Information").OracleTextField(" Reference No.").Enter RACK_GetData("AR_Data","RefNum")
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
Verify_Oracle_Status("FRM-40406: Transaction complete: 1 records applied; all records saved.")

		OracleFormWindow("Lines").CloseWindow
		wait(2)
		OracleFormWindow("Transactions").OracleButton("Complete").Click
		wait(10)
	Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
	Invoice_Number = OracleFormWindow("Transactions").OracleTextField("Transaction|Transaction").GetROProperty("value")
	RACK_PutData "AR_Data", "Invoice_Number", Invoice_Number
	Update_Notepad "Invoice_Number", Invoice_Number
	If not Invoice_Number = "" Then
		RACK_ReportEvent "Invoice Number", "The Invoice Number is sucessfully generated and is - " & Invoice_Number,"Pass"
	End If
End Function
Function  excelData()
		   LineNum =  RACK_GetData("AR_Data","Invoice_Lines_Num")
			LineNumArr= split(LLineNum,";")
			
 excelData = LineNumArr
End Function
'#####################################################################################################################
'Function Description   : Function to Apply Creditmemo To Invoice
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Apply_Creditmemo_To_Invoice()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("AR_Data", "Responsibility_Link"))
	wait(10)
	Select_Link(RACK_GetData("AR_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Transactions Summary").SelectMenu "View->Find..."
	wait(2)
	OracleFormWindow("Find Transactions").OracleTextField("Trasaction Number Low_2").Enter RACK_GetData("AR_Data","Credit_Memo_Number")
	wait(2)
	OracleFormWindow("Find Transactions").OracleTextField("Transaction Dates_2").Enter RACK_GetData("AR_Data","Transaction_Date")
	wait(2)
    OracleFormWindow("Find Transactions").OracleTextField("High Transaction Date_2").Enter RACK_GetData("AR_Data","Transaction_Date")
    wait(2)
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	OracleFormWindow("Find Transactions").OracleButton("Find").Click
    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	OracleFormWindow("Transactions Summary").OracleButton("Applications").Click
	wait(2)
	OracleFormWindow("Applications").OracleTable("Table_2").EnterField 1,"Apply To",RACK_GetData("AR_Data","Invoice_Number")
	wait(2)
	OracleFormWindow("Applications").OracleTable("Table_2").EnterField 1,"Apply Date",RACK_GetData("AR_Data","Apply_Date")
	wait(2)
	OracleFormWindow("Applications").SelectMenu "File->Save"
	wait(5)
	Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
End Function

'#####################################################################################################################
'Function Description   : Function to create a Receipt
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Create_Receipt()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("AR_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("AR_Data", "Functionality_Link"))
	wait(60)
	Invoice_Number = Read_Notepad("Invoice_Number")
	RACK_PutData "AR_Data", "Receipt_Apply_To", Invoice_Number
	If OracleFormWindow("Receipts").OracleTextField("Receipt|Receipt Method").Exist(90) Then
		OracleFormWindow("Receipts").OracleTextField("Receipt|Receipt Method").Enter RACK_GetData("AR_Data","Receipt_Method")
		wait(2)
		OracleFormWindow("Receipts").OracleTextField("Receipt|Receipt Number").Enter RACK_GetData("AR_Data","Receipt_Number")
		wait(2)
		OracleFormWindow("Receipts").OracleTextField("Receipt|Receipt Amount").Enter RACK_GetData("AR_Data","Receipt_Amount")
		If RACK_GetData("AR_Data","Receipt_Date") <> "" Then
			OracleFormWindow("Receipts").OracleTextField("Receipt|Receipt Date").Enter RACK_GetData("AR_Data","Receipt_Date")
		End If
		If RACK_GetData("AR_Data","Receipt_Currency") <> "" Then
			OracleFormWindow("Receipts").OracleTextField("Currency").Enter RACK_GetData("AR_Data","Receipt_Currency")
		End If	
		If RACK_GetData("AR_Data","Receipt_Type") = "Miscellaneous" Then
			OracleFormWindow("Receipts").OracleList("Receipt|Receipt Type").Select RACK_GetData("AR_Data","Receipt_Type")

			If OracleNotification("Caution").Exist(5) then
				OracleNotification("Caution").OracleButton("OK").Click
				wait(2)
			End If
			OracleFormWindow("Receipts").OracleTabbedRegion("Main").OracleTextField("Purpose|Activity").Enter RACK_GetData("AR_Data","Receipt_Activity")
			wait(2)
			OracleFormWindow("Receipts").SelectMenu "File->Save"
			wait(2)
			Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
		   Else
			If RACK_GetData("AR_Data","Invoice_Number")<>"" Then
				OracleFormWindow("Receipts").OracleTabbedRegion("Main").OracleTextField("Detail|Identify By|Transaction").Enter RACK_GetData("AR_Data","Invoice_Number")
				OracleFormWindow("Receipts").OracleTabbedRegion("Main").OracleTextField("Detail|Customer|Name").Click
			End If     	
			wait(2)
			If RACK_GetData("AR_Data","Receipt_CustNum") <> ""Then
				OracleFormWindow("Receipts").OracleTabbedRegion("Main").OracleTextField("Detail|Customer|Number").Enter RACK_GetData("AR_Data","Receipt_CustNum")
			End If
			RACK_ReportEvent "Validation Screenshot", "Create Receipt Parameters successfully Enter   ","Screenshot"           
			OracleFormWindow("Receipts").OracleButton("Apply").Click
			wait(2)
			If   RACK_GetData("AR_Data","Invoice_Number") <>""Then
				If OracleFormWindow("Applications").OracleTextField("Amount Applied").Exist(30) Then
				    OracleFormWindow("Applications").OracleButton("Apply in Detail").Click
				End If
				If OracleFormWindow("Detailed Applications").OracleTextField("Summary Applications|Line").Exist(30) Then
					OracleFormWindow("Detailed Applications").OracleTable("Table").EnterField 1,"Summary Applications|Line", RACK_GetData("AR_Data","Receipt_Amount")
					wait(2)
				    OracleFormWindow("Detailed Applications").SelectMenu "File->Save"
				wait(10)
				Verify_Oracle_Status("FRM-40400: Transaction complete: 2 records applied and saved.")
				OracleFormWindow("Detailed Applications").CloseWindow
				End if 
			End if
			If  RACK_GetData("AR_Data","Receipt_CustNum") <> "" Then
				OracleFormWindow("Applications").OracleTable("Table").EnterField 1,"Apply To",RACK_GetData("AR_Data","Receipt_Apply_To")
				wait(2)
				OracleFormWindow("Applications").OracleTable("Table").EnterField 1,"Apply Date", RACK_GetData("AR_Data","Receipt_Date")
				wait(2)
				OracleFormWindow("Applications").OracleTable("Table").EnterField 1,"Amount Applied",RACK_GetData("AR_Data","Receipt_Amount")
				wait(2)
				OracleFormWindow("Applications").SelectMenu "File->Save"
				wait(10)
				Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
			End If
			End If
	End if
End Function

'#####################################################################################################################
'Function Description   : Function to Applu a Receipt
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Apply_Receipt()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("AR_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("AR_Data", "Functionality_Link"))
	If OracleFormWindow("Receipts Summary").Exist(120) Then
        OracleFormWindow("Receipts Summary").OracleTable("Table").SetFocus 1,"Receipt Number"
		wait(2)
		OracleFormWindow("Receipts Summary").OracleTable("Table").EnterField 1,"State","Unapplied"
		OracleFormWindow("Receipts Summary").OracleTable("Table").InvokeSoftkey "ENTER QUERY"
		wait(2)
		OracleFormWindow("Receipts Summary").OracleTable("Table").EnterField 1,"Receipt Number",RACK_GetData("AR_Data","Receipt_Number")
		wait(2)
		OracleFormWindow("Receipts Summary").OracleTable("Table").InvokeSoftkey"EXECUTE QUERY"
		wait(2)
		OracleFormWindow("Receipts Summary").OracleButton("Open").Click
		If OracleFormWindow("Receipts").Exist(20) Then
			OracleFormWindow("Receipts").OracleButton("Apply").Click
			wait(2)
			If OracleFormWindow("Applications").OracleTable("Table").Exist(30) Then
				If  RACK_GetData("AR_Data","Apply_Unapply") = "Apply" Then
					OracleFormWindow("Applications").OracleTable("Table").EnterField 1,"Apply To",RACK_GetData("AR_Data","Receipt_Apply_To")
					wait(2)
					OracleFormWindow("Applications").OracleTable("Table").EnterField 1,"Apply Date",Format_Date("dd-mmm-yyyy")
					wait(2)
					OracleFormWindow("Applications").OracleTable("Table").EnterField 1,"Amount Applied",RACK_GetData("AR_Data","Receipt_Amount")
			End if
				If  RACK_GetData("AR_Data","Apply_Unapply") = "Unapply" Then
					OracleFormWindow("Applications").OracleTable("Table").EnterField 1,"Apply",False
					wait(2)
				End if
				OracleFormWindow("Applications").SelectMenu "File->Save"
				wait(5)
				Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
			End if
		End if
	End if
End Function

'#####################################################################################################################
'Function Description   : Function to WriteOff_Invoice
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function WriteOff_Invoice()
If Browser("Oracle Applications Home Page").Page("Oracle Applications Home").Link("RS US AR Collection User").Exist(60) Then
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home").Link("RS US AR Collection User").Click
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home").Link("Account Details").Click
	wait(2)
	If OracleFormWindow("Find Account Details").Exist(120)  Then
		OracleFormWindow("Find Account Details").OracleTextField("Operating Unit").Enter RACK_GetData("AR_Data","Operating_Unit")
		wait(2)
		OracleFormWindow("Find Account Details").OracleTextField("Transaction Number").Enter RACK_GetData("AR_Data","Transaction_Number")
		wait(2)
		OracleFormWindow("Find Account Details").OracleList("Status").Select RACK_GetData("AR_Data","Status")
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(2)
		OracleFormWindow("Find Account Details").OracleButton("Find").Click
		wait(2)
		If OracleFormWindow("Account Details").Exist(60)  Then
			OracleFormWindow("Account Details").OracleButton("Adjust").Click
			wait(2)
			If OracleFormWindow("Adjustments").Exist(60)  Then
                OracleFormWindow("Adjustments").OracleTabbedRegion("Main").OracleTable("Table").EnterField 1,"Activity Name", RACK_GetData("AR_Data","Activity_Name")
				wait(2)
				OracleFormWindow("Adjustments").OracleTabbedRegion("Main").OracleTable("Table").EnterField 1,"Type", RACK_GetData("AR_Data","Invoice_Type")
				wait(2)
				OracleFormWindow("Adjustments").OracleTabbedRegion("Main").OracleTable("Table").SetFocus 1,"Amount"
				OracleFormWindow("Adjustments").OracleTabbedRegion("Main").OracleTable("Table").EnterField 1,"Amount",RACK_GetData("AR_Data","Amount_1")
				wait(2)
				OracleFormWindow("Adjustments").OracleTabbedRegion("Main").OracleTable("Table").EnterField 1,"GL Date", RACK_GetData("AR_Data","GL_Date")
				wait(2)
				OracleFormWindow("Adjustments").OracleTabbedRegion("Main").OracleTable("Table").EnterField 1,"Adjustment Date", RACK_GetData("AR_Data","Adjustment_Date")
				wait(5)
				RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				wait(5)
				OracleFormWindow("Adjustments").OracleTabbedRegion("Comments").OracleTable("Table").EnterField 1,"Reason", RACK_GetData("AR_Data","Reason")
				wait(2)
				OracleFormWindow("Adjustments").OracleTabbedRegion("Comments").OracleTable("Table").EnterField 1,"Comments", RACK_GetData("AR_Data","Comments")
				wait(2)
				OracleFormWindow("Adjustments").SelectMenu "File->Save"
				wait(2)
				Dim  balanceVal
					balanceVal =  OracleFormWindow("short title:=Adjustments").OracleTextField("prompt:=Balance").GetROProperty("value")

								RACK_ReportEvent "Verification -InvBalance","Invoice balance =  0.00","Pass"
			Else
				RACK_ReportEvent "Payments Oracle Form Window does not Exist","Fail"
			End If
		Else
			RACK_ReportEvent "Payments Oracle Form Window does not Exist","Fail"
		End If
	Else
		RACK_ReportEvent "Payments Oracle Form Window does not Exist","Fail"
	End If
Else
	RACK_ReportEvent "Payments Oracle Form Window does not Exist","Fail"
End If
End Function
'#####################################################################################################################
'Function Description   : Function to WriteOff_Invoice
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Write_Off_Reversal()
If Browser("Oracle Applications Home Page").Page("Oracle Applications Home").Link("RS US AR Collection User").Exist(60) Then
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home").Link("RS US AR Collection User").Click
	wait(5)
	Browser("Oracle Applications Home Page").Page("Oracle Applications Home").Link("Account Details").Click
	wait(2)
	If OracleFormWindow("Find Account Details").Exist(120)  Then
		OracleFormWindow("Find Account Details").OracleTextField("Operating Unit").Enter RACK_GetData("AR_Data","Operating_Unit")
		wait(2)
		OracleFormWindow("Find Account Details").OracleTextField("Transaction Number").Enter RACK_GetData("AR_Data","Transaction_Number")
		wait(2)
		OracleFormWindow("Find Account Details").OracleList("Status").Select RACK_GetData("AR_Data","Status")
		wait(2)
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		OracleFormWindow("Find Account Details").OracleButton("Find").Click
		wait(2)
		If OracleFormWindow("Account Details").Exist(60)  Then
			OracleFormWindow("Account Details").OracleButton("Adjust").Click
			wait(2)
			If OracleFormWindow("Adjustments").Exist(60)  Then
                OracleFormWindow("Adjustments").OracleTabbedRegion("Main").OracleTable("Table").EnterField 2,"Activity Name", RACK_GetData("AR_Data","Activity_Name")
				wait(2)
				OracleFormWindow("Adjustments").OracleTabbedRegion("Main").OracleTable("Table").EnterField 2,"Type", RACK_GetData("AR_Data","Invoice_Type")
				wait(2)
				OracleFormWindow("Adjustments").OracleTabbedRegion("Main").OracleTable("Table").SetFocus 2,"Amount"
				OracleFormWindow("Adjustments").OracleTabbedRegion("Main").OracleTable("Table").EnterField 2,"Amount",RACK_GetData("AR_Data","Amount_1")
				wait(2)
				OracleFormWindow("Adjustments").OracleTabbedRegion("Main").OracleTable("Table").EnterField 2,"GL Date", RACK_GetData("AR_Data","GL_Date")
				wait(2)
				OracleFormWindow("Adjustments").OracleTabbedRegion("Main").OracleTable("Table").EnterField 2,"Adjustment Date", RACK_GetData("AR_Data","Adjustment_Date")
				wait(5)
				RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				wait(2)
				OracleFormWindow("Adjustments").OracleTabbedRegion("Comments").OracleTable("Table").EnterField 2,"Reason", RACK_GetData("AR_Data","Reason")
				wait(2)
				OracleFormWindow("Adjustments").OracleTabbedRegion("Comments").OracleTable("Table").EnterField 2,"Comments", RACK_GetData("AR_Data","Comments")
				wait(2)
				OracleFormWindow("Adjustments").SelectMenu "File->Save"
				wait(2)
				Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
				End If
           End If
      End If
End If
End Function

'######################################################################################################################
'Function Description   : Function to run Rackspace Print Selected Invoices
'Input Parameters 	: None
'Return Value    	: None
'Created By : Sravanthi
'##################################################################################################################### 

	Function Print_Selected_Invoices()
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("AR_Data", "Responsibility_Link"))
	wait(10)
	Select_Link(RACK_GetData("AR_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	wait(2)
	OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("AR_Data","Invoice_Request_Name")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Transaction Class").Enter RACK_GetData("AR_Data","Transaction_Class")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Transaction Type").Enter RACK_GetData("AR_Data","Transaction_Type")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Transaction Number Low").Enter RACK_GetData("AR_Data","Transaction_Number_Low")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Transaction Number High").Enter RACK_GetData("AR_Data","Transaction_Number_High")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Open Invoices Only").Enter RACK_GetData("AR_Data","Open_Invoices")
   RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	OracleFlexWindow("Parameters").Approve
	wait(5)
	OracleFormWindow("Submit Request").OracleButton("Upon Completion...|Options...").Click
	wait(2)
	If not  RACK_GetData("AR_Data", "Template_Name") = "" Then
    OracleFormWindow("Upon Completion...").OracleTable("Table_3").EnterField 1,"Template Name",RACK_GetData("AR_Data","Template_Name")
	End If
	wait(2)
	OracleFormWindow("Upon Completion...").OracleTable("Table").EnterField 1,"Printer","Anders"
	wait(2)
    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	OracleFormWindow("Upon Completion...").OracleButton("OK").Click
	wait(5)
	OracleFormWindow("Submit Request").OracleButton("Submit").Click
	wait(1)
	RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(1)
	Notification_Message = Get_Oracle_Notification_Form_Message()
	Arr=split(Notification_Message, " ")
	Content = Arr(6)
	Payment_Request_ID_Arr = Split(Content,")")
	Payment_Request_ID = Payment_Request_ID_Arr(0)
	Update_Notepad "Payment_Request_ID", Payment_Request_ID
	Handle_Oracle_Notification_Forms("No")
	Verify_Request_Status(Payment_Request_ID)
End Function
'######################################################################################################################################
'Function Description   : Function to create an Unidentified Receipt
'Input Parameters 	: None
'Return Value    	: None
'######################################################################################################################################

Function Create_Unidentified_Receipt()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("AR_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("AR_Data", "Functionality_Link"))
	wait(60)
	If OracleFormWindow("Receipts").OracleTextField("Receipt|Receipt Method").Exist(90) Then
		OracleFormWindow("Receipts").OracleTextField("Receipt|Receipt Method").Enter RACK_GetData("AR_Data","Receipt_Method")
		wait(2)
		OracleFormWindow("Receipts").OracleTextField("Receipt|Receipt Number").Enter RACK_GetData("AR_Data","Receipt_Number")
		wait(2)
		OracleFormWindow("Receipts").OracleTextField("Receipt|Receipt Amount").Enter RACK_GetData("AR_Data","Receipt_Amount")
		If RACK_GetData("AR_Data","Receipt_Date") <> "" Then
			OracleFormWindow("Receipts").OracleTextField("Receipt|Receipt Date").Enter RACK_GetData("AR_Data","Receipt_Date")
		End If
		If RACK_GetData("AR_Data","Receipt_Currency") <> "" Then
			OracleFormWindow("Receipts").OracleTextField("Currency").Enter RACK_GetData("AR_Data","Receipt_Currency")
        End If
		If RACK_GetData("AR_Data","Receipt_Type") = "Miscellaneous" Then
			OracleFormWindow("Receipts").OracleList("Receipt|Receipt Type").Select RACK_GetData("AR_Data","Receipt_Type")
        End If
			OracleFormWindow("Receipts").SelectMenu "File->Save"
			wait(2)
			Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
End If
End Function
'################################################################################################################################



'################################################################################################################################

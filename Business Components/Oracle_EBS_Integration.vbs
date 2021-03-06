'#####################################################################################################################
'Function Description   : Function to create a receivable Invoice
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Create_Invoices()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Integration_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Integration_Data", "Functionality_Link"))
	wait(60)
	If OracleFormWindow("Transactions").OracleTextField("Transaction|Source").Exist(10) Then
			OracleFormWindow("Transactions").OracleList("Transaction|Class").Select RACK_GetData("Integration_Data","Invoice_Class")
			wait(2)
			OracleFormWindow("Transactions").OracleTextField("Transaction|Type").Enter RACK_GetData("Integration_Data","Invoice_Type")
		If RACK_GetData("Integration_Data","Transaction_Date") <> "" Then
			OracleFormWindow("Transactions").OracleTextField("Transaction|Date").Enter RACK_GetData("Integration_Data","Transaction_Date")			
		End If		
		If RACK_GetData("Integration_Data","Transaction_Currency") <> "" Then
			OracleFormWindow("Transactions").OracleTextField("Transaction|Currency").Enter RACK_GetData("Integration_Data","Transaction_Currency")			
		End If
	wait(2)
		If RACK_GetData("Integration_Data","Invoice_BillTo_Name") <> "" Then
			OracleFormWindow("Transactions").OracleTabbedRegion("Main").OracleTextField("Bill To: Name").Enter RACK_GetData("Integration_Data","Invoice_BillTo_Name")
		else
			OracleFormWindow("Transactions").OracleTabbedRegion("Main").OracleTextField("Bill To: Number").Enter RACK_GetData("Integration_Data","Invoice_BillTo_Number")
		End If
		wait(2)
		If RACK_GetData("Integration_Data","Payment_Term") <> "" Then
			OracleFormWindow("Transactions").OracleTabbedRegion("Main").OracleTextField("Payment Term").Enter RACK_GetData("Integration_Data","Payment_Term")
End If
Wait(2)
'***************Added code to pouplate bank and crdit card details in DFF of invoice to print ***********
If RACK_GetData("Integration_Data","Bank_Ccinfo")= "Yes" Then
OracleFormWindow("Transactions").OracleTextField("Transaction|[").SetFocus
OracleFlexWindow("Transaction Information").OracleTextField("Bank Information").Enter RACK_GetData("Integration_Data","Bank_Information")
OracleFlexWindow("Transaction Information").OracleTextField("Credit Card Information").Enter RACK_GetData("Integration_Data","Credit_Card_Information")
OracleFlexWindow("Transaction Information").Approve
		End if
		wait(2)
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		OracleFormWindow("Transactions").OracleButton("Line Items").Click
		end if
		Dim LineNum,i
		Dim LineDesc,Qty,UtPrice
        Dim LineNumArr,LineDescArr,QtyArr,UtPriceArr
		LineNum =  RACK_GetData("Integration_Data","Invoice_Lines_Num")
		LineNumArr = split(LineNum,";")
        LineDesc = RACK_GetData("Integration_Data","Invoice_Lines_Description")
		LineDescArr= split(LineDesc,";")
		Qty = RACK_GetData("Integration_Data","Invoice_Lines_Quantity")
		QtyArr = split(Qty,";")
		UtPrice = RACK_GetData("Integration_Data","Invoice_Lines_UnitPrice")
		UtPriceArr = split(UtPrice,";")
		LinePrd = RACK_GetData("Integration_Data","Invoice_Lines_Product")
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
				OracleFlexWindow("Invoice Line Information").OracleTextField("Billing Cycle").Enter RACK_GetData("Integration_Data","Invoice_Lines_BillingCycle")
			   OracleFlexWindow("Invoice Line Information").OracleTextField("Product").Enter LinePrdArr(i)
			  If RACK_GetData("Integration_Data","RefNum") <> "" Then
				 OracleFlexWindow("Invoice Line Information").OracleTextField(" Reference No.").Enter RACK_GetData("Integration_Data","RefNum")
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
	RACK_PutData "Integration_Data", "Invoice_Number", Invoice_Number
	If not Invoice_Number = "" Then
		RACK_ReportEvent "Invoice Number", "The Invoice Number is sucessfully generated and is - " & Invoice_Number,"Pass"
	End If
End Function
Function  excelData()
		   LineNum =  RACK_GetData("Integration_Data","Invoice_Lines_Num")
			LineNumArr= split(LLineNum,";")
			
 excelData = LineNumArr
End Function
'######################################################################################################################
'Function Description   : Function to run NL DD Outbound Request Set
'Input Parameters 	: None
'Return Value    	: None
'Created By : Sravanthi
'##################################################################################################################### 
	Function NL_DD_Outbound_Process()
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("Integration_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Integration_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
	wait(3)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	wait(3)
	OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("Integration_Data","Request_Name")
	wait(3)
	OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
	OracleFlexWindow("Parameters").OracleTextField("Currency Code").Enter RACK_GetData("Integration_Data","Currency_Code")
	'OracleFlexWindow("Parameters").OracleTextField("Currency Code").Enter RACK_GetData("Integration_Data","Currency_Code")
	OracleFlexWindow("Parameters").Approve
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
	End Function
'#####################################################################################################################
'Function Description   : Function to run NL DD Inbound Request Set
'Input Parameters 	: None
'Return Value    	: None
'Created BY : Sravanthi
'##################################################################################################################### 
	Function NL_DD_Inbound_Process()
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("Integration_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Integration_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
	wait(3)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("Integration_Data","Request_Name")
	wait(3)
	OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
	OracleFlexWindow("Parameters").OracleTextField("Posting Date").Enter RACK_GetData("Integration_Data","Posting_Date")
	OracleFlexWindow("Parameters").Approve
    wait(3)
	OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 2,"Parameters"
	OracleFlexWindow("Parameters").OracleTextField("Currency Code").Enter RACK_GetData("Integration_Data","Currency_Code")
	wait(3)
	OracleFlexWindow("Parameters").OracleTextField("Application Type").Enter RACK_GetData("Integration_Data","Application_Type")
	OracleFlexWindow("Parameters").Approve
	wait(3)
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
	End Function
 '#####################################################################################################################
'Function Description   : Function to run UK DD Inbound Request Set
'Input Parameters 	: None
'Return Value    	: None
'Created By : Sravanthi
'##################################################################################################################### 
	Function UK_DD_Pears_Process()                   
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("Integration_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Integration_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
	wait(2)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("Integration_Data","Request_Name")
	'OracleListOfValues("Sets").Select "Rackspace UK Direct Debit Outbound Request Set"
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Date From").Enter RACK_GetData("Integration_Data","From_date")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Date To").Enter RACK_GetData("Integration_Data","To_Date")
	OracleFlexWindow("Parameters").Approve
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 2,"Parameters"
	OracleFlexWindow("Parameters").OracleTextField("Filename").Enter RACK_GetData("Integration_Data","File_Name")
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
End Function
'#####################################################################################################################
'Function Description   : Function to run UK DD Inbound Request Set
'Input Parameters 	: None
'Return Value    	: None
'Created By : Sravanthi
'##################################################################################################################### 
	Function UK_DD_Process()	
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("Integration_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Integration_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
	wait(2)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("Integration_Data","Request_Name")
	'OracleListOfValues("Sets").Select "Rackspace UK Direct Debit Outbound Request Set"
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
	'OracleFlexWindow("Parameters").OracleTextField("Payment Method").Enter RACK_GetData("Integration_Data","Payment_Method")
	'OracleFlexWindow("Parameters").Approve
	'OracleFlexWindow("Parameters").OracleTextField("Currency Code").Enter RACK_GetData("Integration_Data","Currency_code"
	OracleFlexWindow("Parameters").Approve
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 2,"Parameters"
	OracleFlexWindow("Parameters").OracleTextField("Filename").Enter RACK_GetData("Integration_Data","File_Name")
	wait(2)
	OracleFlexWindow("Parameters").Approve
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
End Function
'#####################################################################################################################
'Function Description   : Function to attach a Missellaneous Payment
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Misc_File_Upload()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Integration_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Integration_Data", "Functionality_Link"))
	If Parent_desc.WebFile("uploadFile_oafileUpload").Exist(20) Then
        wait(2)
		Parent_desc.WebFile("uploadFile_oafileUpload").Set RACK_GetData("Integration_Data", "Attachement_Path") '"C:\Users\221045\Desktop\sample.txt"
		File_Name = Split(RACK_GetData("Integration_Data", "Attachement_Path"),"\")
		Link_Name = File_Name(ubound(File_Name))
		wait(2)
		Parent_desc.WebButton("Submit").Click
		wait(2)
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
'Function Description   : Function to attach a Missellaneous Payment
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Lockbox_File_Upload()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("Integration_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Integration_Data", "Functionality_Link"))
	If Parent_desc.WebFile("uploadFile_oafileUpload").Exist(20) Then
        wait(2)
		Parent_desc.WebFile("uploadFile_oafileUpload").Set RACK_GetData("Integration_Data", "Attachement_Path") '"C:\Users\221045\Desktop\sample.txt"
		File_Name = Split(RACK_GetData("Integration_Data", "Attachement_Path"),"\")
		Link_Name = File_Name(ubound(File_Name))
		wait(2)
		Parent_desc.WebButton("Submit").Click
		wait(2)
		If Parent_desc.Link("name:="& Link_Name).Exist(20) then
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is uploaded as expected","Pass"
		else
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is not uploaded as expected","Fail"
		end if
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	end if
End Function
''#####################################################################################################################
''Function Description   : Function for Misc payments
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 
Function Misc_Payments()
'set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("Integration_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Integration_Data", "Functionality_Link"))
	If OracleFormWindow("Submit a New Request").Exist(60)  Then
		OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request Set").Exist(60)  Then
			OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("Integration_Data","Request_Name")
			wait(2)
            OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
				OracleFlexWindow("Parameters").OracleTextField("File Name").Enter RACK_GetData("Integration_Data","File_Name")
				OracleFlexWindow("Parameters").Approve
				wait(2)
			else
				RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
			end if
			OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 2,"Parameters"
			OracleFlexWindow("Parameters").Approve
            wait(2)
			OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 3,"Parameters"
				OracleFlexWindow("Parameters").OracleTextField("Currency Code").Enter RACK_GetData("Integration_Data","Currency")
				wait(2)
				OracleFlexWindow("Parameters").OracleTextField("Application Type").Enter "All"
				OracleFlexWindow("Parameters").Approve
				wait(2)
			OracleFormWindow("Submit Request Set").OracleButton("Submit").Click
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			Notification_Message = Get_Oracle_Notification_Form_Message()
			Arr=split(Notification_Message, " ")
			Content = Arr(6)
			Misc_Payment_Request_ID_Arr = Split(Content,")")
			Misc_Payment_Request_ID = Misc_Payment_Request_ID_Arr(0)
			Update_Notepad "Misc_Payment_Request_ID", Misc_Payment_Request_ID
			Handle_Oracle_Notification_Forms("No")
			Verify_Request_Status(Misc_Payment_Request_ID)
		end if
End Function
'#####################################################################################################################
'Function Description   : Function to run LockBox Program
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Run_LockBox_Program()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Lockbox Batch Summary")
	Select_Link(RACK_GetData("Integration_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Integration_Data", "Functionality_Link"))
	If Parent_desc.WebButton("Submit New Batch").Exist(20) then
		Parent_desc.WebButton("Submit New Batch").Click
	end if
	set Child_desc = Browser("Oracle Applications R12").Page("Submit Load And Stage")
	RACK_ReportEvent "Create Customer", "The button Create Customer is clicked","Done"
	If Child_desc.WebEdit("Fndcpprogramdesc").Exist(20) then
		Child_desc.WebEdit("Fndcpprogramdesc").Set RACK_GetData("Integration_Data", "Attachement_Path")
		wait(2)
		Browser("Submit Load And Stage").Page("Submit Load And Stage").WebEdit("Fndcpprogramdesc").Set "Lockbox"
		wait(2)
		Child_desc.WebButton("Next").Click
		wait(2)
		Browser("Submit Load And Stage").Page("Submit Load And Stage_2").WebEdit("N590").Set "Lockbox"
		Child_desc.WebButton("Next").Click
		wait(2)
		Child_desc.WebButton("Submit").Click
		wait(2)
		Message = Child_desc.WebTable("Your request for Rackspace").GetCellData(1,1)
		Array_Message = split (Message, " ")
		LockBox_Request_ID = Array_Message(17)
		Update_Notepad "LockBox_Request_ID", LockBox_Request_ID
		RACK_ReportEvent "LockBox_Request_ID", "LockBox_Request_ID is created sucessfully and is '" & LockBox_Request_ID & "'" ,"Pass"
		RACK_ReportEvent "Validation Screenshot", "Validation Screenshot" ,"Screenshot"
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
End Function
'#####################################################################################################################
'Function Description   : Function to run LockBox Program
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Search_Lockbox_Batch()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Lockbox Batch Summary")
	Select_Link(RACK_GetData("Integration_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Integration_Data", "Functionality_Link"))
	wait(5)
	If not  RACK_GetData("Integration_Data", "GL_Date") = "" Then
		Parent_desc.WebEdit("SearchGLDate").Set RACK_GetData("Integration_Data", "GL_Date")
	End If
	If not  RACK_GetData("Integration_Data", "Request_id") = "" Then
		Parent_desc.WebEdit("SearchConcReqId").Set RACK_GetData("Integration_Data", "Request_id")
	End if
	If not  RACK_GetData("Integration_Data", "Btach_Number") = "" Then
		Parent_desc.WebEdit("SearchBatchNumber").Set RACK_GetData("Integration_Data", "Btach_Number")
	End if
	If Parent_desc.WebButton("Go").Exist(20) then
		Parent_desc.WebButton("Go").Click
	'end if
	wait(2)
	Browser("Oracle Applications Home Page").Page("Lockbox Batch Summary").Link("20140922052208").Click
    wait(2)
    Browser("Oracle Applications Home Page").Page("Lockbox Batch Detail").Link("Ready To Process").Click
	wait(2)
    Browser("Oracle Applications Home Page").Page("Lockbox Receipt Detail").Link("Select All").Click
	wait(2)
    Browser("Oracle Applications Home Page").Page("Lockbox Receipt Detail").WebButton("Verified").Click
	wait(2)
		RACK_ReportEvent "Validation Screenshot", "Validation Screenshot" ,"Screenshot"
	'else
		'RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
End Function
'#####################################################################################################################
'Function Description   : Function to run LockBox Program
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 
Function Lockbox_Post_Batch()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Lockbox Batch Summary")
	Select_Link(RACK_GetData("Integration_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Integration_Data", "Functionality_Link"))
	wait(5)
	If not  RACK_GetData("Integration_Data", "GL_Date") = "" Then
		Parent_desc.WebEdit("SearchGLDate").Set RACK_GetData("Integration_Data", "GL_Date")
	End If
	If not  RACK_GetData("Integration_Data", "Request_id") = "" Then
		Parent_desc.WebEdit("SearchConcReqId").Set RACK_GetData("Integration_Data", "Request_id")
	End if
	If not  RACK_GetData("Integration_Data", "Btach_Number") = "" Then
		Parent_desc.WebEdit("SearchBatchNumber").Set RACK_GetData("Integration_Data", "Btach_Number")
	End if
	If Parent_desc.WebButton("Go").Exist(20) then
		Parent_desc.WebButton("Go").Click
	'end if
 wait(2)
   Browser("Oracle Applications Home Page").Page("Lockbox Batch Summary").Link("20140922052208").Click
wait(2)
   Browser("Oracle Applications Home Page").Page("Lockbox Batch Detail").WebButton("Post Batch").Click
wait(2)
   Browser("Oracle Applications Home Page").Page("Warning").WebButton("Submit Batch").Click
wait(2)
		RACK_ReportEvent "Validation Screenshot", "Validation Screenshot" ,"Screenshot"
	'else
		'RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	End If
End Function
''#####################################################################################################################
''Function Description   : Function to run Single Request
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 
Function Run_Single_Request()
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("Integration_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Integration_Data", "Functionality_Link"))
	If OracleFormWindow("Submit a New Request").Exist(60)  Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(60)  Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Integration_Data", "Request_Name")
			wait(2)
			If  OracleFlexWindow("Parameters").Exist(10) Then
				OracleFlexWindow("Parameters").OracleTextField("Currency Code").Enter RACK_GetData("Integration_Data", "Currency")
				wait(2)
				OracleFlexWindow("Parameters").OracleTextField("Application Type").Enter RACK_GetData("Integration_Data","Application_Type")
                RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
				OracleFlexWindow("Parameters").OracleButton("OK").Click
				wait(2)
			End If
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
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
	End if
End Function
''#####################################################################################################################
''Function Description   : Function for CC_Global_Gateway
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 
Function CC_Global_Gateway_Outbound()
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("Integration_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Integration_Data", "Functionality_Link"))
	wait(60)
	If OracleFormWindow("Submit a New Request").Exist(60)  Then
		OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		wait(2)
			OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("Integration_Data","Request_Name")
			wait(2)
			'If OracleFlexWindow("Parameters").Exist(10)  Then
               OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
               OracleFlexWindow("Parameters").Approve
			   wait(2)
                 OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 2,"Parameters"
                  OracleFlexWindow("Parameters").OracleTextField("Currency").Enter "USD"
                    OracleFlexWindow("Parameters").Approve
				  wait(2)
                        OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 3,"Parameters"
                          OracleFlexWindow("Parameters").Approve
		End if
			OracleFormWindow("Submit Request Set").OracleButton("Submit").Click
			RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			Notification_Message = Get_Oracle_Notification_Form_Message()
			Arr=split(Notification_Message, " ")
			Content = Arr(6)
			CC_Global_Gateway_Request_ID_Arr = Split(Content,")")
			CC_Global_Gateway_Request_ID = CC_Global_Gateway_Request_ID_Arr(0)
			Update_Notepad "CC_Global_Gateway_Request_ID", CC_Global_Gateway_Request_ID
			Handle_Oracle_Notification_Forms("No")
			Verify_Request_Status(CC_Global_Gateway_Request_ID)
RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
	'End if
End Function
''#####################################################################################################################
''Function Description   : Function for CC_Global_Gateway
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 
Function CC_Global_Gateway_Inbound()
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("Integration_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Integration_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("Integration_Data","Request_Name")
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 1,"Parameters"
	OracleFlexWindow("Parameters").OracleTextField("Posting Date").Enter RACK_GetData("Integration_Data","Posting_Date")
	OracleFlexWindow("Parameters").Approve
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 2,"Parameters"
	OracleFlexWindow("Parameters").OracleTextField("Currency Code").Enter RACK_GetData("Integration_Data","Currency")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Application Type").Enter RACK_GetData("Integration_Data","Application_Type")
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
End Function
'#####################################################################################################################
''Function Description   : Function to run Journal Import
''Input Parameters 	: None
''Return Value    	: None
''##################################################################################################################### 
Function Blackline_Outbound_Set()
'set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("Integration_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("Integration_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request_2").Select "Request Set"
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTextField("Run this Request...|Request").Enter RACK_GetData("Integration_Data","Request_Name")
	wait(2)
	OracleFormWindow("Submit Request Set").OracleTable("Table").SetFocus 2,"Parameters"
	if not RACK_GetData("Integration_Data","Outbound_Directory") = "" then
	OracleFlexWindow("Parameters").OracleTextField("Outbound Directory").Enter RACK_GetData("Integration_Data","Outbound_Directory")
	End If
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
	End Function
'########################################################################################################################################
'
'
'
'########################################################################################################################################

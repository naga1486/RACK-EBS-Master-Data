
'#####################################################################################################################
'CPU Patch Testing
'Function Description   : Function to create contract through Invoice Worksheet
'Input Parameters 	: None
'Return Value    	: None
'Created By : Avvaru Nagarjuna
'##################################################################################################################### 

Function Load_Transactions_AR()
 set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("SC_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("SC_Data", "Functionality_Link"))
	If OracleFormWindow("Manual Invoice Worksheet").Exist(120) Then
		Dim TranType,TransTypeArr,CustNum,CustNumArr,PeriodDate,PeriodDateArr,Memoline,MemoLineArr,GLPrdname,GLPrdnameArr,Qty,QtyArr
		Dim UtPrice,UtPriceArr,i
		TranType = RACK_GetData("SC_Data", "Transaction_Type")
		TransTypeArr = split(TranType,";")
		PeriodDate = RACK_GetData("SC_Data", "Period_Date")
		PeriodDateArr = split(PeriodDate,";")
		CustNum = RACK_GetData("SC_Data", "Customer_Number")
		CustNumArr = split(CustNum,";")
		Memoline = RACK_GetData("SC_Data", "Memo_Line")
		MemoLineArr = split(Memoline,";")
		GLPrdname = RACK_GetData("SC_Data", "GL_Product_Name")
		GLPrdnameArr = split(GLPrdname,";")
		Qty = RACK_GetData("SC_Data", "Quantity")
		QtyArr = split(Qty,";")
		UtPrice = RACK_GetData("SC_Data", "Unit_Price")
		UtPriceArr = split(UtPrice,";")
        DCount = RACK_GetData("SC_Data", "DCount")
        For  i = 0 to DCount-1
			If i=0 Then
              OracleFormWindow("Manual Invoice Worksheet").OracleTable("Table").EnterField 1,"Transaction Type",TransTypeArr(i)
             wait(2)
             OracleFormWindow("Manual Invoice Worksheet").OracleTable("Table").EnterField 1,"Period Date",PeriodDateArr(i)
			wait(2)
			OracleFormWindow("Manual Invoice Worksheet").OracleTable("Table").EnterField 1,"Customer Number",CustNumArr(i)
			wait(2)
			OracleFormWindow("Manual Invoice Worksheet").OracleTable("Table").EnterField 1,"Memo Line",MemoLineArr(i)
			wait(2)
		OracleFormWindow("Manual Invoice Worksheet").OracleTable("Table").SetFocus 1,"GL Product Name"
			wait(3)
    	OracleFormWindow("Manual Invoice Worksheet").OracleTable("Table_2").EnterField 1,"GL Product Name",GLPrdnameArr(i)
			wait(2)
			OracleFormWindow("Manual Invoice Worksheet").OracleTable("Table_2").EnterField 1,"Quantity",QtyArr(i)
			wait(2)
		OracleFormWindow("Manual Invoice Worksheet").OracleTable("Table_2").EnterField 1,"UOM", "Ea"
			wait(2)
			OracleFormWindow("Manual Invoice Worksheet").OracleTable("Table_2").EnterField 1,"Unit Price",UtPriceArr(i)
            			wait(2)
							End If
			OracleFormWindow("Manual Invoice Worksheet").SelectMenu "File->Save"
			GLPrdname=OracleFormWindow("Manual Invoice Worksheet").OracleTable("Table_2").GetFieldValue (1,"GL Product Name")
			If GLPrdname = "" Then
				OracleFormWindow("Manual Invoice Worksheet").OracleTable("Table_2").EnterField 1,"GL Product Name",GLPrdnameArr(i) '"MANAGED"
			End If
		Next
		OracleFormWindow("Manual Invoice Worksheet").SelectMenu "File->Save"
		wait(5)
		Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
End if
End Function

'#####################################################################################################################
'Function Description   : Function to run Rackspace Billing Creation Process
'Input Parameters 	: None
'Return Value    	: None
'Created by : Avvaru Nagarjuna
'##################################################################################################################### 

Function Billing_Creation_Process()
	set WshShell = CreateObject("WScript.Shell")
    set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("SC_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("SC_Data", "Functionality_Link"))
	wait(90)
	If OracleFormWindow("Submit a New Request").Exist(10)  Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(10)  Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("SC_Data", "Request_Name")
		wait(2)
	If  OracleFlexWindow("Parameters").Exist(10) Then
	     OracleFlexWindow("Parameters").OracleTextField("Data Source to Create").Enter RACK_GetData("SC_Data","Data_Source")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Billing Run Type").Enter RACK_GetData("SC_Data","Bill_Run_Type")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Billing Period").Enter RACK_GetData("SC_Data","Bill_Period")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Invoice Date").Enter RACK_GetData("SC_Data","Invoice_Date")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Data Processing Options").Enter RACK_GetData("SC_Data","Process_Mode")
	wait(2)
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
End If
End Function

'#####################################################################################################################
'Function Description   : Function to Run Rackspace Usage Data Import Process
'Input Parameters 	: None
'Return Value    	: None
'Created by : Avvaru Nagarjuna
'##################################################################################################################### 

Function Usage_Data_Import_Process
	set WshShell = CreateObject("WScript.Shell")
    set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("SC_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("SC_Data", "Functionality_Link"))
	wait(30)
	If OracleFormWindow("Submit a New Request").Exist(60)  Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(60)  Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("SC_Data", "Request_Name")
			wait(2)
			If  OracleFlexWindow("Parameters").Exist(10) Then
				OracleFlexWindow("Parameters").OracleTextField("Usage Data Type").Enter RACK_GetData("SC_Data", "Data_Type")
				wait(2)
				OracleFlexWindow("Parameters").OracleTextField("Billing Period").Enter RACK_GetData("SC_Data", "Period")
				wait(2)
				OracleFlexWindow("Parameters").OracleTextField("Usage Data Input File").Enter RACK_GetData("SC_Data", "p_file_name")
				wait(2)
				OracleFlexWindow("Parameters").OracleTextField("Processing Options").Enter RACK_GetData("SC_Data", "Process_Mode")
				RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
                wait(2)
				OracleFlexWindow("Parameters").Approve
			wait(2)
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
End If
End If
End Function

'#####################################################################################################################
'Function Description   : Function to  Upload Usage Csv File
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Upload_Usage_Files()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("SC_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("SC_Data", "Functionality_Link"))
	wait(10)
	If Parent_desc.WebFile("uploadFile_oafileUpload").Exist(20) Then
        wait(2)
		Parent_desc.WebFile("uploadFile_oafileUpload").Set RACK_GetData("SC_Data", "File_Attachement_Path") '"C:\Users\221045\Desktop\sample.txt"
		File_Name = Split(RACK_GetData("SC_Data", "File_Attachement_Path"),"\")
		Link_Name = File_Name(ubound(File_Name))
		wait(2)
		Parent_desc.WebButton("Submit").Click
		wait(5)
		If Parent_desc.Link("name:="& Link_Name).Exist(20) then
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is uploaded as expected","Pass"
		else
			RACK_ReportEvent "File Upload", "The File '" & Link_Name & "' is not uploaded as expected","Fail"
		end if
	else
		RACK_ReportEvent "Page doesnt Open", "Page doesnt Open" ,"Fail"
End if
End Function

''#####################################################################################################################
''Function Description   : Function for Run Rackspace Invoice Worksheet Interface to staging program
''Input Parameters 	: None
''Return Value    	: None
'Created by : Avvaru Nagarjuna
''##################################################################################################################### 

Function IWI_Staging_Program()
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("SC_Data", "Responsibility_Link"))
	wait(10)
	Select_Link(RACK_GetData("SC_Data", "Functionality_Link"))
	If OracleFormWindow("Submit a New Request").Exist(90)  Then
		'OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request").Select "Single Request"
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(60)  Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("SC_Data", "Request_Name")
			wait(2)
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
			wait(2)
			OracleNotification("Caution").Approve
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
End if
End Function
''#####################################################################################################################
''Function Description   : Function for Run Rackspace Mass Team and GL Account Update program
''Input Parameters 	: None
''Return Value    	: None
'Created by : Avvaru Nagarjuna
''##################################################################################################################### 

Function Mass_GL_Account_Program()
'set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("SC_Data", "Responsibility_Link"))
	wait(10)
	Select_Link(RACK_GetData("SC_Data", "Functionality_Link"))
	If OracleFormWindow("Submit a New Request").Exist(90)  Then
		'OracleFormWindow("Submit a New Request").OracleRadioGroup("Single Request").Select "Single Request"
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(60)  Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("SC_Data", "Request_Name")
			wait(2)
			OracleFlexWindow("Parameters").Approve
			wait(2)
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
End if
End Function

''#####################################################################################################################
''Function Description   : Function for Run Rackspace Device Name Change
''Input Parameters 	: None
''Return Value    	: None
'Created by : Avvaru Nagarjuna
''##################################################################################################################### 

Function Change_Device_Name()
	set WshShell = CreateObject("WScript.Shell")
set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("SC_Data", "Responsibility_Link"))
	wait(3)
	Select_Link(RACK_GetData("SC_Data", "Functionality_Link"))
	wait(30)
	If OracleFormWindow("Submit a New Request").Exist(60)  Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
		If OracleFormWindow("Submit Request").Exist(60)  Then
			OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("SC_Data", "Request_Name")
			wait(2)
			If  OracleFlexWindow("Parameters").Exist(10) Then
				 OracleFlexWindow("Parameters").OracleTextField("Since").Enter RACK_GetData("SC_Data", "Period_Date")
				wait(2)
				OracleFlexWindow("Parameters").OracleButton("OK").Click
				wait(2)
			End If
			OracleFormWindow("Submit Request").OracleButton("Submit").Click
			RACK_ReportEvent "Validation Screenshot", "Rackspace Device Name Change Parameters successfully Enter   ","Screenshot"
		End If
		Notification_Message = Get_Oracle_Notification_Form_Message()
		Arr=split(Notification_Message, " ")
		Content = Arr(6)
		Payment_Request_ID_Arr = Split(Content,")")
		Payment_Request_ID = Payment_Request_ID_Arr(0)
		Update_Notepad "Payment_Request_ID", Payment_Request_ID
		Handle_Oracle_Notification_Forms("No")
		Verify_Request_Status(Payment_Request_ID)
End If
End Function

'#####################################################################################################################
'Function Description   : Verify_Contracts
'Input Parameters 	: Noner
'Return Value    	: Nonemg
'Created by : Avvaru Nagarjuna
'##################################################################################################################### 

Function Verify_Contracts()
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Suppliers")
	Select_Link(RACK_GetData("SC_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("SC_Data", "Functionality_Link"))
	wait(60)
	If RACK_GetData("SC_Data","Transaction_Number") <> "" Then
		OracleFormWindow("Contract Find").OracleTextField("Transaction No.").Enter RACK_GetData("SC_Data", "Transaction_Number")
		End If
		If RACK_GetData("SC_Data","Date_Received_From") <> "" Then
		OracleFormWindow("Contract Find").OracleTextField("Date Received From").Enter RACK_GetData("SC_Data", "Date_Received_From")
		End If
		If RACK_GetData("SC_Data","Date_Received_To") <> "" Then
		OracleFormWindow("Contract Find").OracleTextField("Date Received To").Enter RACK_GetData("SC_Data", "Date_Received_To")
		End If
		OracleFormWindow("Contract Find").OracleButton("Find").Click
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(10)
		OracleFormWindow("Contract Find").OracleButton("Open").Click
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	    wait(5)
		OracleFormWindow("Contracts").OracleTabbedRegion("Contracts").OracleButton("Device").Click
		wait(5)
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(5)
End Function
'#####################################################################################################################
'Function Description   : Function to create a Auto Invoice
'Input Parameters 	: None
'Return Value    	: None
'##################################################################################################################### 

Function Auto_Invoice_Master_Program()
	set WshShell = CreateObject("WScript.Shell")
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Customers")
	Select_Link(RACK_GetData("SC_Data", "Responsibility_Link"))
	wait(10)
	Select_Link(RACK_GetData("SC_Data", "Functionality_Link"))
	If OracleFormWindow("Submit a New Request").OracleButton("OK").Exist(120) Then
		OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	end if
	If  OracleFormWindow("Run AutoInvoice").OracleTextField("Run this Request...|Name").Exist(30) Then
		OracleFormWindow("Run AutoInvoice").OracleTextField("Run this Request...|Name").Enter RACK_GetData("SC_Data","Request_Name")
		OracleFormWindow("Run AutoInvoice").OracleButton("Submit").Click
		If  OracleNotification("Error").Exist(5) Then
			OracleNotification("Error").OracleButton("OK").Click
		End If
	End If
	If  OracleFlexWindow("Parameters").OracleTextField("Number of Instances").Exist(30)Then
		OracleFlexWindow("Parameters").OracleTextField("Invoice Source").Enter RACK_GetData("SC_Data","Invoice_Source")
		wait(1)
		OracleFlexWindow("Parameters").OracleTextField("Default Date").Enter RACK_GetData("SC_Data","Invoice_Default_Date")
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		OracleFlexWindow("Parameters").OracleButton("OK").Click
		OracleFormWindow("Run AutoInvoice").OracleButton("Submit").Click
		RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	End If
		Notification_Message = Get_Oracle_Notification_Form_Message()
			Arr=split(Notification_Message, " ")
			Content = Arr(6)
			Misc_Payment_Request_ID_Arr = Split(Content,")")
			Misc_Payment_Request_ID = Misc_Payment_Request_ID_Arr(0)
			Update_Notepad "Misc_Payment_Request_ID", Misc_Payment_Request_ID
			Handle_Oracle_Notification_Forms("No")
			Verify_Request_Status(Misc_Payment_Request_ID)
End Function
'######################################################################################################################
'Function Description   : Function to run Rackspace Print Selected Invoices
'Input Parameters 	: None
'Return Value    	: None
'Created By : Sravanthi
'##################################################################################################################### 
	Function Print_Selected_Invoice()
	set Parent_desc = Browser("Oracle Applications Home Page").Page("Oracle Applications Home")
	Select_Link(RACK_GetData("SC_Data", "Responsibility_Link"))
	wait(10)
	Select_Link(RACK_GetData("SC_Data", "Functionality_Link"))
	wait(60)
	OracleFormWindow("Submit a New Request").OracleButton("OK").Click
	wait(2)
	OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("SC_Data","Request_Name")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Transaction Class").Enter RACK_GetData("SC_Data","Transaction_Class")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Transaction Type").Enter RACK_GetData("SC_Data","Transaction_Type")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Transaction Number Low").Enter RACK_GetData("SC_Data","Transaction_Number_Low")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Transaction Number High").Enter RACK_GetData("SC_Data","Transaction_Number_High")
'	wait(2)
'	OracleFlexWindow("Parameters").OracleTextField("Print Date Low").Enter RACK_GetData("SC_Data","PrintDate_Low")
'	wait(2)
'	OracleFlexWindow("Parameters").OracleTextField("Print Date High").Enter RACK_GetData("SC_Data","PrintDate_High")
	'wait(2)
	'OracleFlexWindow("Parameters").OracleTextField("Customer").Enter RACK_GetData("SC_Data","Customer_Name")
	wait(2)
	OracleFlexWindow("Parameters").OracleTextField("Open Invoices Only").Enter RACK_GetData("SC_Data","Open_Invoices")
   RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
	wait(2)
	OracleFlexWindow("Parameters").Approve
	wait(5)
	OracleFormWindow("Submit Request").OracleButton("Upon Completion...|Options...").Click
	wait(2)
	'OracleFormWindow("Upon Completion...").OracleTable("Table").OpenDialog 1,"Template Name"
	If not  RACK_GetData("SC_Data", "Template_Name") = "" Then
    OracleFormWindow("Upon Completion...").OracleTable("Table_3").EnterField 1,"Template Name",RACK_GetData("SC_Data","Template_Name")
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
'End If
'End If
	End Function

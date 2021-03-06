
'#####################################################################################################################
'Function Description   : Function to clear transactions
'Input Parameters 	: None
'Return Value    	: None
'Created by : Sravanthi
'##################################################################################################################### 
Function Manually_Clear_Transactions()
	Select_Link(RACK_GetData("CM_Data", "Responsibility_Link"))
	wait(5)
	 Select_Link(RACK_GetData("CM_Data", "Functionality_Link"))
	If OracleFormWindow("Find Transactions").Exist(120)  Then
		OracleFormWindow("Find Transactions").OracleTabbedRegion("Bank").OracleTextField("Account Number").SetFocus
		OracleFormWindow("Find Transactions").OracleTabbedRegion("Bank").OracleTextField("Account Number").Enter RACK_GetData("CM_Data", "Account_Number") 
		If not RACK_GetData("CM_Data", "MiscCheckBox") = "Yes" Then
			OracleFormWindow("Find Transactions").OracleCheckbox("Miscellaneous").Clear
		End If
		If not RACK_GetData("CM_Data", "CMCheckBox") = "Yes" Then
			OracleFormWindow("Find Transactions").OracleCheckbox("Cash Management Cashflow").Clear
		End If
		If not RACK_GetData("CM_Data", "APCheckBox") = "Yes" Then
			OracleFormWindow("Find Transactions").OracleCheckbox("AP Payment").Clear
		End If
		If not RACK_GetData("CM_Data", "ARCheckBox") = "Yes" Then
			OracleFormWindow("Find Transactions").OracleCheckbox("AR Receipt").Clear
		End If
'		OracleFormWindow("Find Transactions").OracleTabbedRegion("Bank").OracleTextField("Account Number").SetFocus
'		OracleFormWindow("Find Transactions").OracleTabbedRegion("Bank").OracleTextField("Account Number").Enter RACK_GetData("CM_Data", "Account_Number") 
		If not RACK_GetData("CM_Data", "TransNumLow") = "" Then
			OracleFormWindow("Find Transactions").OracleTabbedRegion("Transaction").OracleTextField("Numbers").Enter RACK_GetData("CM_Data", "TransNumLow")
			OracleFormWindow("Find Transactions").OracleTabbedRegion("Transaction").OracleTextField("High Number").Enter RACK_GetData("CM_Data", "TransNumHigh")
		End If
		If not RACK_GetData("CM_Data", "TrasDateLow") = "" Then
			OracleFormWindow("Find Transactions").OracleTabbedRegion("Transaction").OracleTextField("Dates").Enter RACK_GetData("CM_Data", "TrasDateLow")
			OracleFormWindow("Find Transactions").OracleTabbedRegion("Transaction").OracleTextField("High Date").Enter RACK_GetData("CM_Data", "TrasDateHigh")
		End If
		If not RACK_GetData("CM_Data", "TransBatchLow") = "" Then
			OracleFormWindow("Find Transactions").OracleTabbedRegion("Transaction").OracleTextField("Batch Names").Enter RACK_GetData("CM_Data", "TransBatchLow")
			OracleFormWindow("Find Transactions").OracleTabbedRegion("Transaction").OracleTextField("High Batch Name").Enter RACK_GetData("CM_Data", "TransBatchHigh")
		End If
		If not RACK_GetData("CM_Data", "Currency") = "" Then
			OracleFormWindow("Find Transactions").OracleTabbedRegion("Transaction").OracleTextField("Currency").Enter RACK_GetData("CM_Data", "Currency")
		End If
		If not RACK_GetData("CM_Data", "Status") = "" Then
			OracleFormWindow("Find Transactions").OracleTabbedRegion("Transaction").OracleTextField("Status").Enter RACK_GetData("CM_Data", "Status")
		End If
		OracleFormWindow("Find Transactions").OracleButton("Find").Click
		wait(10)
		OracleFormWindow("Clear Transactions").OracleTabbedRegion("Transaction").OracleTable("Table").EnterField 1,"Select Record",true
		OracleFormWindow("Clear Transactions").OracleTabbedRegion("Transaction").OracleTable("Table").EnterField 2,"Select Record",true
		OracleFormWindow("Clear Transactions").OracleTabbedRegion("Transaction").OracleTable("Table").EnterField 3,"Select Record",true
        RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
		wait(5)
		OracleFormWindow("Clear Transactions").OracleButton("Clear Transaction").Click
		'ClearTransNumber = OracleFormWindow("Clear Transactions").OracleTabbedRegion("Transaction").OracleTable("Table").GetFieldValue( 1,"Number")
		'RACK_PutData "CM_Data", "ClearTransNumber", ClearTransNumber
		'OracleFormWindow("Clear Transactions").SelectMenu "File->Save"
		'wait(2)
		Verify_Oracle_Status("FRM-40400: Transaction complete: 3 records applied and saved.")
'		wait(5)
'		OracleFormWindow("Clear Transactions").CloseWindow
'		wait(5)
'		if OracleNotification("Forms").Exist(10) then
'			OracleNotification("Forms").Approve
		End if
'	End IF
End Function

'#####################################################################################################################
'Function Description   : Function to Create a new Bank Statement
'Input Parameters 	: None
'Return Value    	: None
'Created by : Sravanthi
'##################################################################################################################### 
Function Create_New_Bank_Statement()
	Select_Link(RACK_GetData("CM_Data", "Responsibility_Link"))
	wait(5)
    Browser("Oracle Applications Home Page").Page("Oracle Applications Home_27").Link("Bank Statements and Reconciliation").Click
	If OracleFormWindow("Find Bank Statements").Exist(120)  Then
		OracleFormWindow("Find Bank Statements").OracleTextField("Account Number").SetFocus
		OracleFormWindow("Find Bank Statements").OracleButton("New").Click
		wait(5)
		OracleFormWindow("Bank Statement").OracleTextField("Account Number").SetFocus
		OracleFormWindow("Bank Statement").OracleTextField("Account Number").Enter RACK_GetData("CM_Data", "Account_Number")
		wait(2)
		OracleFormWindow("Bank Statement").OracleTextField("Date").Enter RACK_GetData("CM_Data", "Date")
		OracleNotification("Error_2").Approve
			wait(2)
			OracleFormWindow("Bank Statement").OracleTextField("Control Totals|Opening").Enter RACK_GetData("CM_Data", "Amount")
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
             wait(2)
			OracleFormWindow("Bank Statement").OracleButton("Available").Click
			wait(2)
			OracleNotification("Error_2").Approve
			wait(2)
            OracleFormWindow("Find Transactions").OracleTabbedRegion("Transaction").OracleTextField("Dates").Enter RACK_GetData("CM_Data", "TrasDateLow")
			wait(2)
            OracleFormWindow("Find Transactions").OracleTabbedRegion("Transaction").OracleTextField("High Date").Enter RACK_GetData("CM_Data", "TrasDateHigh")
			wait(2)
			If not RACK_GetData("CM_Data", "MiscCheckBox") = "Yes" Then
				OracleFormWindow("Find Transactions").OracleCheckbox("Miscellaneous").Clear
			End If
			If not RACK_GetData("CM_Data", "CMCheckBox") = "Yes" Then
				OracleFormWindow("Find Transactions").OracleCheckbox("Cash Management Cashflow").Clear
			End If
			If not RACK_GetData("CM_Data", "APCheckBox") = "Yes" Then
				OracleFormWindow("Find Transactions").OracleCheckbox("AP Payment").Clear
			End If
			If not RACK_GetData("CM_Data", "ARCheckBox") = "Yes" Then
				OracleFormWindow("Find Transactions").OracleCheckbox("AR Receipt").Clear
			End If
			If not RACK_GetData("CM_Data", "JournalCheckBox") = "Yes" Then
				OracleFormWindow("Find Transactions").OracleCheckbox("Journal").Clear
			End If
			If not RACK_GetData("CM_Data", "PayrollCheckBox") = "Yes" Then
				OracleFormWindow("Find Transactions").OracleCheckbox("Payroll").Clear
			End If
			If not RACK_GetData("CM_Data", "EFTCheckBox") = "Yes" Then
				OracleFormWindow("Find Transactions").OracleCheckbox("Payroll EFT").Clear
			End If
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
			OracleFormWindow("Find Transactions").OracleButton("Find").Click
			wait(5)
			OracleFormWindow("Available Transactions").OracleTabbedRegion("Transaction").OracleTable("Table").EnterField 1,"Select Record",true
			ReconNumber = OracleFormWindow("Available Transactions").OracleTabbedRegion("Transaction").OracleTable("Table").GetFieldValue( 1,"Number")
			RACK_PutData "CM_Data", "ReconNumber", ReconNumber
			If OracleNotification("Note").Exist(10)  Then
				OracleNotification("Note").Approve
			End if
            RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"	 
			OracleFormWindow("Available Transactions").OracleButton("Reconcile").Click
			OracleNotification("Decision").Approve
			OracleFormWindow("Available Transactions").SelectMenu "File->Save"
			OracleFormWindow("Available Transactions").CloseWindow
		wait(5)	
		OracleFormWindow("Bank Statement").SelectMenu "File->Save"
		OracleFormWindow("Bank Statement").OracleCheckbox("Complete").Select
		OracleFormWindow("Bank Statement").SelectMenu "File->Save"
		OracleNotification("Decision").Approve	
		Verify_Oracle_Status("FRM-40400: Transaction complete: 1 records applied and saved.")
If not  RACK_GetData("CM_Data", "LineNum") = "" Then
			OracleFormWindow("Bank Statement").OracleButton("Lines").Click
			RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"	
	End if
	End if  
End Function
'#####################################################################################################################
'Function Description   : Function to View Availabl Transactions 
'Input Parameters 	: None
'Return Value    	: None
'Created by : Sravanthi
'##################################################################################################################### 
Function View_Availabl_Transactions ()
	Select_Link(RACK_GetData("CM_Data", "Responsibility_Link"))
wait(5)
  Select_Link(RACK_GetData("CM_Data", "Functionality_Link"))
If OracleFormWindow("Find Transactions").Exist(120)  Then
If not RACK_GetData("CM_Data", "Account_Number") = "" Then
             OracleFormWindow("Find Transactions").OracleTabbedRegion("Bank").OracleTextField("Account Number").Enter RACK_GetData("CM_Data", "Account_Number") 
             End If
If not RACK_GetData("CM_Data", "Account_Name") = "" Then
              OracleFormWindow("Find Transactions").OracleTabbedRegion("Bank").OracleTextField("Account Name").Enter RACK_GetData("CM_Data", "Account_Name") 
              End If
If not RACK_GetData("CM_Data", "MiscCheckBox") = "Yes" Then
			   OracleFormWindow("Find Transactions").OracleCheckbox("Miscellaneous").Clear
			    End If
If not RACK_GetData("CM_Data", "CMCheckBox") = "Yes" Then
		     OracleFormWindow("Find Transactions").OracleCheckbox("Cash Management Cashflow").Clear
			  End If
If not RACK_GetData("CM_Data", "APCheckBox") = "Yes" Then
		   OracleFormWindow("Find Transactions").OracleCheckbox("AP Payment").Clear
		   End If
If not RACK_GetData("CM_Data", "ARCheckBox") = "Yes" Then
			OracleFormWindow("Find Transactions").OracleCheckbox("AR Receipt").Clear
			 End If
If not RACK_GetData("CM_Data", "JournalCheckBox") = "Yes" Then
		    OracleFormWindow("Find Transactions").OracleCheckbox("Journal").Clear
			End If
If not RACK_GetData("CM_Data", "PayrollCheckBox") = "Yes" Then
			OracleFormWindow("Find Transactions").OracleCheckbox("Payroll").Clear
			 End If
If not RACK_GetData("CM_Data", "EFTCheckBox") = "Yes" Then
			OracleFormWindow("Find Transactions").OracleCheckbox("Payroll EFT").Clear
			 End If
             RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
             wait(2)
             OracleFormWindow("Find Transactions").OracleButton("Find").Click
             RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
wait(10)
End If
End Function
'#####################################################################################################################
'Function Description   : Function to View Bank Statements
'Input Parameters 	: None
'Return Value    	: None
'Created by : Sravanthi
'#####################################################################################################################
Function View_Bank_Statements()
Select_Link(RACK_GetData("CM_Data", "Responsibility_Link"))
wait(5)
Browser("Oracle Applications Home Page").Page("Oracle Applications Home_27").Link("Bank Statements and Reconciliation").Click
If OracleFormWindow("Find Bank Statements").Exist(120)  Then
OracleFormWindow("Find Bank Statements").OracleTextField("Account Number").SetFocus
OracleFormWindow("Find Bank Statements").OracleTextField("Account Number").Enter RACK_GetData("CM_Data", "Account_Number")
wait(5)
OracleFormWindow("Find Bank Statements").OracleButton("Find").Click
wait(5)
OracleFormWindow("View Bank Statement Reconciliation").SelectMenu "View->Find..."
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
wait(5)
OracleFormWindow("Find Bank Statements").CloseWindow
wait(5)
OracleFormWindow("View Bank Statement Reconciliation").OracleButton("Review").Click
wait(5)
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
wait(5)
OracleFormWindow("View Bank Statement").OracleButton("Lines").Click
wait(5)
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
wait(5)
OracleFormWindow("View Bank Statement Lines").OracleButton("Available").Click
wait(5)
RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
End If
End Function
'#####################################################################################################################
'Function Description   : Function to View Receipts
'Input Parameters 	: None
'Return Value    	: None
'Created by : Sravanthi
'#####################################################################################################################
Function View_Receipts()
	Select_Link(RACK_GetData("CM_Data", "Responsibility_Link"))
	wait(5)
	Select_Link(RACK_GetData("CM_Data", "Functionality_Link"))
If OracleFormWindow("Receipts").Exist(120)  Then
    OracleFormWindow("Receipts").SelectMenu "View->Query By Example->Enter"
If not RACK_GetData("CM_Data", "Receipt_Method") = "" Then
   OracleFormWindow("Receipts").OracleTextField("Receipt|Receipt Method").Enter RACK_GetData("CM_Data", "Receipt_Method")
   End If
   If not RACK_GetData("CM_Data", "Receipt_Number") = "" Then
   OracleFormWindow("Receipts").OracleTextField("Receipt|Receipt Number").Enter RACK_GetData("CM_Data", "Receipt_Number")
   End If
   RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
wait(2)
   OracleFormWindow("Receipts").SelectMenu "View->Query By Example->Run"
   wait(5)
   RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
   OracleFormWindow("Receipts").OracleButton("Apply").Click
   wait(5)
    RACK_ReportEvent "Validation Screenshot", "Screenshot ","Screenshot"
End If
End Function

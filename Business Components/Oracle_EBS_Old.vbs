Randomize 
Sub Login_Old() 
'		If Browser("Login").Page("Oracle Applications Home_2").Link("Logout").Exist(1) Then
'			 Browser("Login").Page("Oracle Applications Home_2").Link("Logout").Click
'		end if
'        If 	Browser("Login").Exist(1) Then
'			Browser("Login").CloseAllTabs()
'		end if
'		If Browser("Oracle Applications R12").Exist(1) Then
'			Browser("Oracle Applications R12").CloseAllTabs()
'		End If		
	'While Browser("CreationTime:=0").Exist
		'Browser("CreationTime:=0").Close
	'Wend
'	While Browser("CreationTime:=1").Exist
'		Browser("CreationTime:=1").Close
'	Wend

   'SystemUtil.Run "iexplore","https://patch.orapi.rackspace.com/"
	SystemUtil.Run "iexplore",RACK_GetData("Common_Data", "Application_URL" ),"","",3
	If   Browser("Login").Page("Login").WebEdit("usernameField").Exist(120) Then
        Browser("Login").Page("Login").WebEdit("usernameField").Set RACK_GetData("Common_Data", "username")
		RACK_ReportEvent "Login ", "User name "&RACK_GetData("Common_Data", "username" )& " is entered ","Done"
 		Browser("Login").Page("Login").WebEdit("passwordField").SetSecure RACK_GetData("Common_Data", "password" )
		Browser("Login").Page("Login").WebButton("Login").Click
		RACK_ReportEvent "Login ", "Login button is clicked","Done"
		If Browser("Login").Page("Oracle Applications Home_2").Link("Logout").Exist(20) Then
				Reporter.ReportEvent micPass, "Login ", "Successfully logined"
				RACK_ReportEvent "Login ", "Successfully logined","Done"
		Else 
				RACK_ReportEvent  "Login ", "Home Page does not exist", "Fail"
		End If
	Else
			RACK_ReportEvent  "Login ", "Login Failed - Username text box does not exist", "Fail"
	End If
End Sub
Sub Logout()	
	If Browser("Login").Page("Oracle Applications Home_2").Link("Logout").Exist(20) Then
			 Browser("Login").Page("Oracle Applications Home_2").Link("Logout").Click
			RACK_ReportEvent "Logout ", "Logout link is clicked","Done"
			 RACK_ReportEvent  "Logout" ,"Sucessfully exited the application","Done"
			If 		Browser("Login").Exist(5) Then
						Browser("Login").CloseAllTabs()
						If Browser("Oracle Applications R12").Exist(5) Then
							Browser("Oracle Applications R12").CloseAllTabs()
						End If		
			End If
	Else
			 RACK_ReportEvent  "Logout" ,"Logout Unsucessfull", "Fail"
	End If
End Sub
Sub Create_Requisition()
   Select Case RACK_GetData("Login_Data", "Type" )
	Case "Non Catalogue"
		If Browser("Login").Page("Oracle Applications Home").Link("RS US iProcurement Requestor").Exist(20) Then
			Browser("Login").Page("Oracle Applications Home").Link("RS US iProcurement Requestor").Click
			If Browser("Login").Page("Oracle iProcurement: Shop").Link("Non-Catalog Request").Exist(20) Then
				Browser("Login").Page("Oracle iProcurement: Shop").Link("Non-Catalog Request").Click
				If Browser("Login").Page("Oracle iProcurement: Shop").WebEdit("ItemDescription").Exist(20) Then
					Browser("Login").Page("Oracle iProcurement: Shop").WebEdit("ItemDescription").Set "sample"
					Browser("Login").Page("Oracle iProcurement: Shop").WebEdit("Category").Set "48U cabinet"
					Browser("Login").Page("Oracle iProcurement: Shop").WebEdit("Quantity").Set "1"
					Browser("Login").Page("Oracle iProcurement: Shop").WebEdit("UnitPrice").Set "10"
					If Browser("Login").Page("Oracle iProcurement: Shop").WebButton("Add to Cart_2").Exist(10) Then
						Browser("Login").Page("Oracle iProcurement: Shop").WebButton("Add to Cart_2").Click
						'Browser("Login").Page("Oracle iProcurement: Shop").WebButton("Add to Cart").Click
						If Browser("Login").Page("Oracle iProcurement: Shop").WebButton("View Cart and Checkout").Exist(20) Then
							Browser("Login").Page("Oracle iProcurement: Shop").WebButton("View Cart and Checkout").Click
							If Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Checkout").Exist(20) Then
								Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Checkout").Click
								If Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Edit Lines").Exist(20) Then
									Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Edit Lines").Click
									Browser("Login").Page("Oracle iProcurement: Checkout").WebCheckBox("DeliveryLinesAdvTable:selected:0").Set "ON"
									If Browser("Login").Page("Oracle iProcurement: Checkout").Link("Attachments").Exist(20) Then
										Browser("Login").Page("Oracle iProcurement: Checkout").Link("Attachments").Click
										If Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Add Attachment...").Exist(20) Then
											Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Add Attachment...").Click
											Browser("Login").Page("Oracle iProcurement: Add").WebEdit("FileName").Set "sample"
											Browser("Login").Page("Oracle iProcurement: Add").WebEdit("AkDescription").Set "sample"
											If Browser("Login").Page("Oracle iProcurement: Add").WebFile("FileInput_oafileUpload").Exist(20) Then
												wait(2)
                                                Browser("Login").Page("Oracle iProcurement: Add").WebFile("FileInput_oafileUpload").Set "C:\Users\221045\Desktop\sample.txt"
												If Browser("Login").Page("Oracle iProcurement: Add").WebButton("Apply").Exist(20) Then
													Browser("Login").Page("Oracle iProcurement: Add").WebButton("Apply").Click
													If Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Apply").Exist(20) Then
														Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Apply").Click
														If Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Next").Exist(20) Then
															Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Next").Click
															If Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Manage Approvals").Exist(20) Then
																Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Next").Click
																If Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Submit").Exist(20) Then
																	Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Submit").Click
																	strReqID=Browser("Login").Page("Confirmation").WebElement("Requisition 44506").GetROProperty("innertext")
																	arrReqID= Split(strReqID," ")
																	'MsgBox(Trim(arrReqID(1)))
																	Environment.Value("RequisitionID") = Trim(arrReqID(1))
																	RACK_ReportEvent "Create Requisition", "Requesition ID "& Environment.Value("RequisitionID") &" Created sucessfully","Pass"
																Else
																	RACK_ReportEvent "Create Requisition", "Submit button does not exist","Fail"
																End If
															Else
																RACK_ReportEvent "Create Requisition", "Manage Approvals button does not exist","Fail"
															End If
														Else
															RACK_ReportEvent "Create Requisition", "Next button does not exist","Fail"
														End If
													Else
														RACK_ReportEvent "Create Requisition", "Apply button does not exist","Fail"
													End If
												Else
													RACK_ReportEvent "Create Requisition", "Apply button does not exist","Fail"
												End If
											Else
												RACK_ReportEvent "Create Requisition", "File Input Upload does not exist","Fail"
											End If
										Else
											RACK_ReportEvent "Create Requisition", "Add Attachment button does not exist","Fail"
										End If
									Else
										RACK_ReportEvent "Create Requisition", "Attachment button does not exist","Fail"
									End If
								Else
									RACK_ReportEvent "Create Requisition", "Edit lines button does not exist","Fail"
								End If
							Else
								RACK_ReportEvent "Create Requisition", "Checkout button does not exist","Fail"
							End If
						Else
							RACK_ReportEvent "Create Requisition", "View Cart and Checkout button does not exist","Fail"
						End If
					Else
						RACK_ReportEvent "Create Requisition", "Add to Cart button does not exist","Fail"
					End If
				Else
					RACK_ReportEvent "Create Requisition", "Item description web edit does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Create Requisition", "Non-Catalog Request link does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Create Requisition", "RS US iProcurement Requestor link does not exist","Fail"
		End If															
   Case "catalogue"
			If Browser("Login").Page("Oracle Applications Home").Link("RS US iProcurement Requestor").Exist(20) Then		
					Browser("Login").Page("Oracle Applications Home").Link("RS US iProcurement Requestor").Click
					If Browser("Login").Page("Oracle iProcurement: Shop").Link("Rackspace Catalog").Exist(20) Then
						Browser("Login").Page("Oracle iProcurement: Shop").Link("Rackspace Catalog").Click
						If Browser("Login").Page("Oracle iProcurement: Shop").Link("Rackspace - Inventory").Exist(10) Then
							Browser("Login").Page("Oracle iProcurement: Shop").Link("Rackspace - Inventory").Click
							If Browser("Login").Page("Oracle iProcurement: Shop").Link("HARD DISK").Exist(10) Then
								Browser("Login").Page("Oracle iProcurement: Shop").Link("HARD DISK").Click
								If Browser("Login").Page("Oracle iProcurement: Shop").WebButton("Add to Cart").Exist(10) Then
									Browser("Login").Page("Oracle iProcurement: Shop").WebButton("Add to Cart").Click
									If  Browser("Login").Page("Oracle iProcurement: Shop").WebButton("View Cart and Checkout").Exist(20) Then
										Browser("Login").Page("Oracle iProcurement: Shop").WebButton("View Cart and Checkout").Click
											If  Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Checkout").Exist(40) Then
												Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Checkout").Click
												If  Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Next").Exist(20) Then
													Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Next").Click
													If  Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Manage Approvals").Exist (20) Then					
															Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Next").Click
															If Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Submit").Exist(20)  Then
																Browser("Login").Page("Oracle iProcurement: Checkout").WebButton("Submit").Click
																strReqID=Browser("Login").Page("Confirmation").WebElement("Requisition 44506").GetROProperty("innertext")
																arrReqID= Split(strReqID," ")
																'MsgBox(Trim(arrReqID(1)))
																Environment.Value("RequisitionID") = Trim(arrReqID(1))
																 RACK_ReportEvent "Create Requisition", "Requesition ID "& Environment.Value("RequisitionID") &" Created sucessfully","Pass"
															 Else
																RACK_ReportEvent "Create Requisition", "Submit button does not exist","Fail"
															End If
													Else 
															RACK_ReportEvent "Create Requisition", "Manage approvals button does not exist","Fail"
												End If
										Else
												RACK_ReportEvent "Create Requisition", "Next button does not exist","Fail"
										End If
									Else
										RACK_ReportEvent "Create Requisition", "Checkout button does not exist","Fail"
									End If
								Else
									RACK_ReportEvent "Create Requisition", "View cart and checkout button does not exist","Fail"
								End If
							Else
									RACK_ReportEvent "Create Requisition", "Add to cart button does not exist","Fail"
							End If
						Else
							RACK_ReportEvent "Create Requisition", "Hard disk link does not exist","Fail"
						End If
					Else
							RACK_ReportEvent "Create Requisition", "Rackspace inventory link does not exist","Fail"
					End If
			Else
					RACK_ReportEvent "Create Requisition", "Rackspace Catalog link does not exist","Fail"
			End If
	Else
			RACK_ReportEvent "Create Requisition", "RS US iProcurement Requestor link does not exist","Fail"
	End If
   End Select
End Sub
Sub Create_PO()
   If Browser("Login").Page("Oracle Applications Home").Link("RS US Purchasing Staff").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home").Link("RS US Purchasing Staff").Click
		wait(5)
        If Browser("Login").Page("Oracle Applications Home_2").Link("AutoCreate").Exist(20) Then		
			Browser("Login").Page("Oracle Applications Home_2").Link("AutoCreate").Click
			Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
			Browser("Login").Page("Oracle Applications Home_2").Sync
			If OracleFormWindow("Find Requisition Lines").OracleButton("Clear").Exist(120) Then
				OracleFormWindow("Find Requisition Lines").OracleButton("Clear").Click
				wait(5)
				OracleFormWindow("Find Requisition Lines").OracleTextField("Requisition").Enter Environment.Value("RequisitionID")
				OracleFormWindow("Find Requisition Lines").OracleButton("Find").Click
				If OracleFormWindow("AutoCreate Documents").OracleButton("Automatic").Exist(20) Then
					OracleFormWindow("AutoCreate Documents").OracleButton("Automatic").Click
					If OracleFormWindow("New Document").OracleButton("Create").Exist(20) Then
						OracleFormWindow("New Document").OracleButton("Create").Click
						If OracleFormWindow("AutoCreate to Purchase").OracleTextField("Supplier").Exist(20) Then
							PO_ID = OracleFormWindow("AutoCreate to Purchase").OracleTextField("PO, Rev").GetROProperty("value")
						'	MsgBox(PO_ID)
							RACK_ReportEvent "Create PO", "PO "&PO_ID&" Created succecssfully","Pass"
							OracleFormWindow("AutoCreate to Purchase").OracleTextField("Supplier").Enter "03 WORLD, LLC"
							OracleFormWindow("AutoCreate to Purchase").OracleButton("Approve...").Click
							If OracleFormWindow("Approve Document").OracleButton("OK").Exist(20) Then
								 OracleFormWindow("Approve Document").OracleButton("OK").Click
								 status = OracleFormWindow("AutoCreate to Purchase").OracleTextField("Status").GetROProperty("value")
								 RACK_ReportEvent  " PO status", "PO - " & status, "Pass"
							Else
								RACK_ReportEvent  " PO", "OK button does not exist", "Fail"
							End If
						Else
								RACK_ReportEvent  " PO", "Supplier text field does not exist", "Fail"
						End If
					Else
						RACK_ReportEvent  " PO", "create button does not exist", "Fail"
					End If
				Else
					RACK_ReportEvent  " PO", "Automatic button does not exist", "Fail"
				End If
			Else
				RACK_ReportEvent  " PO", "Clear button does not exist", "Fail"
			End If	
		Else
			RACK_ReportEvent  " PO", "AutoCreate link does not exist", "Fail"
		End If
	Else
		RACK_ReportEvent  " PO", "RS US Purchasing Staff link does not exist", "Fail"
   End If
End Sub
Sub Approve_Requisition()
   If  Browser("Login").Page("Oracle Applications Home").Link("RS US Purchasing Approver").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home").Link("RS US Purchasing Approver").Click
		If Browser("Login").Page("Oracle Applications Home_2").Link("Notifications").Exist(20) Then
			Browser("Login").Page("Oracle Applications Home_2").Link("Notifications").Click
			If  Browser("Login").Page("Oracle Workflow: Notifications").Link("Select All").Exist(20) Then
				Browser("Login").Page("Oracle Workflow: Notifications").Link("Select All").Click
            	If  Browser("Login").Page("Oracle Workflow: Notifications").WebButton("Open").Exist(40) Then
					Browser("Login").Page("Oracle Workflow: Notifications").WebButton("Open").Click
					If Browser("Login").Page("Notification Details").WebButton("Approve").Exist(20)  Then
						Browser("Login").Page("Notification Details").WebButton("Approve").Click
						wait(2)
							RACK_ReportEvent "Approve_Requisition", "Sucessfully approved", "Pass"  					
					Else
							RACK_ReportEvent "Approve_Requisition", "Approval button does not exist", "Fail"
					End If
				   Else
						   RACK_ReportEvent "Approve_Requisition", "Open button does not exist", "Fail"
				End If
			Else
						RACK_ReportEvent "Approve_Requisition", "Select All Link does not exist", "Fail"
			End If
		Else
						RACK_ReportEvent "Approve_Requisition", "Notification Link does not exist", "Fail"
		End If
		Else
					RACK_ReportEvent "Approve_Requisition", "RS US Purchasing Approver Link does not exist","Fail"
   End If
End Sub 
Sub Subinventory()
If Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory User,").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory User,").Click
	wait(5)
	Browser("Login").Page("Oracle Applications Home_2").Link("Subinventory Transfer").Click
	Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
	Browser("Login").Page("Oracle Applications Home_2").Sync
	If OracleListOfValues("Organizations").Exist(120) Then
		OracleListOfValues("Organizations").Select RACK_GetData("Login_Data", "Organization")
		If OracleFormWindow("Subinventory Transfers").OracleTextField("Transaction|Type").Exist(20) Then
					OracleFormWindow("Subinventory Transfers").OracleTextField("Transaction|Type").Enter RACK_GetData("Login_Data", "Transaction")
					OracleFormWindow("Subinventory Transfers").OracleCheckbox("Transaction|Serial-Triggered").Select
					If OracleFormWindow("Subinventory Transfers").OracleButton("Transaction Lines").Exist(20) then
							OracleFormWindow("Subinventory Transfers").OracleButton("Transaction Lines").Click										
							OracleFormWindow("Subinventory Transfer").OracleTextField("Item").Enter  RACK_GetData("Login_Data", "item")	
							OracleFormWindow("Subinventory Transfer").OracleTextField("Subinventory").Enter RACK_GetData("Login_Data", "Subinventory")
							OracleFormWindow("Subinventory Transfer").OracleTextField("To Subinv").SetFocus
							OracleFormWindow("Subinventory Transfer").OracleTextField("To Subinv").Enter RACK_GetData("Login_Data", "ToSubinventory")
							OracleFormWindow("Subinventory Transfer").OracleTextField("To Locator").Enter RACK_GetData("Login_Data", "Locator")
							OracleFormWindow("Subinventory Transfer").OracleTextField("Quantity").SetFocus
							OracleFormWindow("Subinventory Transfer").OracleTextField("Quantity").Enter RACK_GetData("Login_Data", "Qty")
							If OracleFormWindow("Subinventory Transfer").OracleButton("Lot / Serial").Exist(20) Then
								OracleFormWindow("Subinventory Transfer").OracleButton("Lot / Serial").Click                            						
									If  OracleFormWindow("Serial Entry").OracleTextField("Start Serial Number").Exist(10) Then
											OracleFormWindow("Serial Entry").OracleTextField("Start Serial Number").OpenDialog
											OracleListOfValues("Serial Numbers").Find "%"
											OracleListOfValues("Serial Numbers").Select RACK_GetData("Login_Data","serialnumber")  							
											OracleFormWindow("Serial Entry").OracleButton("Done").Click
											OracleFormWindow("Subinventory Transfer").SelectMenu "File->Save"
											RACK_ReportEvent  "Subinventory transfer" , " Sucessful", "Pass"			
									Else
										 RACK_ReportEvent  "Subinventory transfer" , " Start Serial Number Textfield does not exist", "Fail"	
									End If
							Else
										 RACK_ReportEvent  "Subinventory transfer" , " Lot/Serial button does not exist","Fail"	
						    End if 	
					Else
							RACK_ReportEvent  "Subinventory transfer" , " Transaction Lines button does not exist", "Fail"	
					End If
			Else
							RACK_ReportEvent  "Subinventory transfer" , " Transaction|Type text field does not exist", "Fail"	
           		End If
		Else
							RACK_ReportEvent  "Subinventory transfer" , " Organizations List of value does not exist", "Fail"	
			End If
	Else
					RACK_ReportEvent  "Subinventory transfer" , " RS US Inventory User Link  does not exist", "Fail"	
	End If
End Sub
Sub MaterialTransaction()
   If Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory User,").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory User,").Click
	If Browser("Login").Page("Oracle Applications Home_2").Link("Material Transactions").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home_2").Link("Material Transactions").Click
        Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
		Browser("Login").Page("Oracle Applications Home_2").Sync
		If  OracleListOfValues("Organizations").Exist(120) Then
					OracleListOfValues("Organizations").Select RACK_GetData("Login_Data", "Organization")
					If OracleFormWindow("Find Material Transactions").OracleTextField("Item").Exist(50) Then
						OracleFormWindow("Find Material Transactions").OracleTextField("Item").Enter  RACK_GetData("Login_Data", "item")	
						OracleFormWindow("Find Material Transactions").OracleButton("Find").Click
							If OracleFormWindow("Material Transactions").OracleTabbedRegion("Transaction Type").Exist(50) Then
									OracleFormWindow("Material Transactions").OracleTabbedRegion("Transaction Type").Click
									transaction =  OracleFormWindow("Material Transactions").OracleTabbedRegion("Transaction Type").OracleTextField("Transaction Type").GetROProperty("value")
									'msgBox(transaction)
									 RACK_ReportEvent  "Material transaction" , "Transaction type is " & 	transaction  , "Pass"
							Else
									RACK_ReportEvent  "Material Transaction" , "Transaction Type Tabbed Region does not exist", "Fail"	
							End If
					Else
									RACK_ReportEvent  "Material Transaction" , "Item Textfield  does not exist", "Fail"	
					End If
			Else
							RACK_ReportEvent  "Material Transaction" , "Organization List of Values  does not exist", "Fail"	
			End If
		Else
							RACK_ReportEvent  "Material Transaction" , "Material Transactions Link does not exist", "Fail"	
    	End If
	Else
			RACK_ReportEvent  "Material Transaction" , "RS US Inventory User Link does not exist", "Fail"	
	End If
End Sub
Sub Report()
Select Case RACK_GetData("Login_Data","ReportNavigation")
	Case "Payables"
		If Browser("Login").Page("Oracle Applications Home").Link("RS US Payables Manager").Exist(20) Then
			Browser("Login").Page("Oracle Applications Home").Link("RS US Payables Manager").Click
			If Browser("Login").Page("Oracle Applications Home_2").Link("Run_2").Exist(20) Then
				Browser("Login").Page("Oracle Applications Home_2").Link("Run_2").Click
				Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
				Browser("Login").Page("Oracle Applications Home_2").Sync
				If OracleFormWindow("Submit a New Request").OracleButton("OK").Exist(120) Then
					OracleFormWindow("Submit a New Request").OracleButton("OK").Click
					OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Login_Data","ReportName")
					If OracleFlexWindow("Parameters").Exist(20) Then
						OracleFlexWindow("Parameters").Approve
						OracleFormWindow("Submit Request").OracleButton("Submit").Click						
						arrmessage=OracleNotification("Decision").GetROProperty("message")
						arrrequestID = Split (arrmessage,"=")
						requestID = Split(arrrequestID(1), ")")
						'MsgBox(Trim(requestID(0)))
						OracleNotification("Decision").Decline				
						OracleFormWindow("Navigator").SelectMenu "View->Requests"
						OracleFormWindow("Find Requests").OracleRadioGroup("My Completed Requests").Select "Specific Requests"						
						OracleFormWindow("Find Requests").OracleTextField("Request ID").Enter Trim(requestID(0))
						OracleFormWindow("Find Requests").OracleButton("Find").Click
						counter =0
						Do
							wait(5)
							OracleFormWindow("Requests").OracleButton("Refresh Data").Click
							counter= counter+1				
						Loop Until (counter >RACK_GetData("Login_Data","ReportExeTime")) OR (OracleFormWindow("Requests").OracleTextField("Phase").GetROProperty("Value") = "Completed" AND OracleFormWindow("Requests").OracleTextField("Status").GetROProperty("Value") = "Normal" )
						If counter > RACK_GetData("Login_Data","ReportExeTime") Then
							 RACK_ReportEvent  "Report " ,"Request ID " &Trim(requestID(0)) & " Not completed . Check  the process manually", "Fail"
						Else
							OracleFormWindow("Requests").OracleButton("View Output").Click
							wait(2)
							RACK_ReportEvent   "Process" &Trim(requestID(0)) , "Sucessfully completed", "Pass"
							If Browser("Browser").Exist(5) Then
								wait(2)
									Browser("Browser").CloseAllTabs()
							End If
						End If
					Else
						RACK_ReportEvent  "Report","Parameters flex window does not exist","Fail"
					End If
				Else
					RACK_ReportEvent  "Report","OK button does not exist","Fail"
				End If
			Else
				RACK_ReportEvent  "Report","Run link does not exist","Fail"
			End If
		Else
			RACK_ReportEvent  "Report","RS US Payables Manager link does not exist","Fail"
		End If					
	Case "Asset"
		If Browser("Login").Page("Oracle Applications Home").Link("RS US Fixed Assets Manager").Exist(20)  Then
			Browser("Login").Page("Oracle Applications Home").Link("RS US Fixed Assets Manager").Click
			If Browser("Login").Page("Oracle Applications Home_2").Link("Run").Exist(20) Then
				Browser("Login").Page("Oracle Applications Home_2").Link("Run").Click				
				Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
				Browser("Login").Page("Oracle Applications Home_2").Sync
				If OracleFormWindow("Submit a New Request").OracleButton("OK").Exist(100) Then
					OracleFormWindow("Submit a New Request").OracleButton("OK").Click
					If OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Exist(20) Then
						OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Login_Data","ReportName")
                        If OracleFormWindow("Submit Request").OracleButton("Submit").Exist(20) Then
							OracleFormWindow("Submit Request").OracleButton("Submit").Click
							If  OracleNotification("Caution").Exist(10) Then
								OracleNotification("Caution").Approve
							End If  
							If OracleNotification("Decision").Exist(20)  Then
                                   arrmessage=OracleNotification("Decision").GetROProperty("message")
									arrrequestID = Split (arrmessage,"=")
									requestID = Split(arrrequestID(1), ")")
									'MsgBox(Trim(requestID(0)))
									OracleNotification("Decision").Decline				
									OracleFormWindow("Navigator").SelectMenu "View->Requests"
									OracleFormWindow("Find Requests").OracleRadioGroup("My Completed Requests").Select "Specific Requests"
									
									OracleFormWindow("Find Requests").OracleTextField("Request ID").Enter Trim(requestID(0))
									OracleFormWindow("Find Requests").OracleButton("Find").Click
									counter =0
											Do
										wait(5)
										OracleFormWindow("Requests").OracleButton("Refresh Data").Click
										counter= counter+1				
									Loop Until (counter >RACK_GetData("Login_Data","ReportExeTime")) OR (OracleFormWindow("Requests").OracleTextField("Phase").GetROProperty("Value") = "Completed" AND OracleFormWindow("Requests").OracleTextField("Status").GetROProperty("Value") = "Normal" )
										If counter > RACK_GetData("Login_Data","ReportExeTime") Then
											 RACK_ReportEvent  "Report " ,"Request ID " &Trim(requestID(0)) & " Not completed . Check  the process manually", "Fail"
										Else
													OracleFormWindow("Requests").OracleButton("View Output").Click
													wait(2)
													RACK_ReportEvent   "Process" &Trim(requestID(0)) , "Sucessfully completed", "Pass"
													If Browser("Browser").Exist(5) Then
														wait(2)
															Browser("Browser").CloseAllTabs()
													End If
										End If									
								Else
									RACK_ReportEvent  "Report","Decision does not exist","Fail"
								End If							
						Else
							RACK_ReportEvent  "Report","Submit submit does not exist","Fail"
						End If
					Else
						RACK_ReportEvent  "Report","Run this Request...|Nametext field does not exist","Fail"
					End If
				Else
					RACK_ReportEvent  "Report","OK button does not exist","Fail"
				End If
			Else
				RACK_ReportEvent  "Report","Run link does not exist","Fail"
			End If
		Else
			RACK_ReportEvent  "Report","RS US Fixed Assets Manager link does not exist","Fail"
		End If
			
	Case "Inventory"
	If Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory User,").Exist(20) Then
			Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory User,").Click
			If Browser("Login").Page("Oracle Applications Home_2").Link("Transactions").Exist(20) Then
				Browser("Login").Page("Oracle Applications Home_2").Link("Transactions").Click
				Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
				Browser("Login").Page("Oracle Applications Home_2").Sync
				If OracleListOfValues("Organizations").Exist(120) Then
					OracleListOfValues("Organizations").Select RACK_GetData("Login_Data","Organization")
					OracleFormWindow("Submit a New Request").OracleButton("OK").Click
					Select Case RACK_GetData("Login_Data","ReportName")
						Case "Rackspace On Hand by Serial Number Report"
							OracleFormWindow("Transaction Reports").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Login_Data","ReportName")
							OracleFlexWindow("Parameters").OracleTextField("Sub Inventory").SetFocus
							OracleFlexWindow("Parameters").OracleTextField("Sub Inventory").OpenDialog
							OracleListOfValues("Sub Inventory").Select "In-Service"
							OracleFlexWindow("Parameters").Approve
							OracleFormWindow("Transaction Reports").OracleButton("Submit").Click
						
						Case "Rackspace Serial Number Query"
							OracleFormWindow("Transaction Reports").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Login_Data","ReportName")
							OracleFlexWindow("Parameters").Approve
							OracleFormWindow("Transaction Reports").OracleButton("Submit").Click
		
						Case "rackspace inventory transaction"
							OracleFormWindow("Transaction Reports").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Login_Data","ReportName")
							OracleFlexWindow("Parameters").OracleTextField("Start Date (Ex: 01-JAN-08)").Enter RACK_GetData("Login_Data","StartDate")
							OracleFlexWindow("Parameters").OracleTextField("End Date (Ex: 01-JAN-08)").Enter RACK_GetData("Login_Data","EndDate")
							OracleFlexWindow("Parameters").OracleTextField("Category Set").Enter "RS Item Categories"
							OracleFlexWindow("RS Item Categories").OracleButton("Cancel").Click
							OracleFlexWindow("Parameters").Approve
							OracleFormWindow("Transaction Reports").OracleButton("Submit").Click
	
						Case "Rackspace Build Subassembly Report"
							OracleFormWindow("Transaction Reports").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Login_Data","ReportName")
							OracleFlexWindow("Parameters").Approve
							OracleFormWindow("Transaction Reports").OracleButton("Submit").Click
					End Select
					If OracleNotification("Decision").Exist(20) Then
						arrmessage=OracleNotification("Decision").GetROProperty("message")
						arrrequestID = Split (arrmessage,"=")
						requestID = Split(arrrequestID(1), ")")
						'MsgBox(Trim(requestID(0)))
						OracleNotification("Decision").Decline				
						OracleFormWindow("Navigator").SelectMenu "View->Requests"
						OracleFormWindow("Find Requests").OracleRadioGroup("My Completed Requests").Select "Specific Requests"
						
						OracleFormWindow("Find Requests").OracleTextField("Request ID").Enter Trim(requestID(0))
						OracleFormWindow("Find Requests").OracleButton("Find").Click
						counter =0
						Do
							wait(5)
							OracleFormWindow("Requests").OracleButton("Refresh Data").Click
							counter= counter+1				
						Loop Until (counter >RACK_GetData("Login_Data","ReportExeTime")) OR (OracleFormWindow("Requests").OracleTextField("Phase").GetROProperty("Value") = "Completed" AND OracleFormWindow("Requests").OracleTextField("Status").GetROProperty("Value") = "Normal" )
							If counter > RACK_GetData("Login_Data","ReportExeTime") Then
								 RACK_ReportEvent  "Report " ,"Request ID " &Trim(requestID(0)) & " Not completed . Check  the process manually", "Fail"
							Else
								OracleFormWindow("Requests").OracleButton("View Output").Click
								wait(2)
								RACK_ReportEvent   "Process" &Trim(requestID(0)) , "Sucessfully completed", "Pass"
								If Browser("Browser").Exist(5) Then
									wait(2)
										Browser("Browser").CloseAllTabs()
								End If
							End If
					Else 
							RACK_ReportEvent  "Report","Decision does not exist","Fail"
					End If
				Else
					RACK_ReportEvent  "Report","Organizations list does not exist","Fail"
				End If
			Else
				RACK_ReportEvent  "Report","Transactions link does not exist","Fail"
			End If
	Else
		RACK_ReportEvent  "Report","RS US Inventory User link does not exist","Fail"
	End If
	Case "PayablesClerk"
		If Browser("Login").Page("Oracle Applications Home").Link("RS US Payables Clerk").Exist(20) Then
			Browser("Login").Page("Oracle Applications Home").Link("RS US Payables Clerk").Click
			If Browser("Login").Page("Oracle Applications Home_2").Link("Run_2").Exist(20) Then
				Browser("Login").Page("Oracle Applications Home_2").Link("Run_2").Click
				Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
				Browser("Login").Page("Oracle Applications Home_2").Sync
				If OracleFormWindow("Submit a New Request").OracleButton("OK").Exist(120) Then
					OracleFormWindow("Submit a New Request").OracleButton("OK").Click
					Select Case RACK_GetData("Login_Data","ReportName")
						Case "Create Accounting"							
							OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Login_Data","ReportName")
							OracleFlexWindow("Parameters").OracleTextField("Ledger").Enter RACK_GetData("Login_Data","Ledger") 							
							OracleFlexWindow("Parameters").OracleTextField("End Date").Enter RACK_GetData("Login_Data","EndDate")
							OracleFlexWindow("Parameters").OracleTextField("Mode").Enter RACK_GetData("Login_Data","Mode")
							OracleFlexWindow("Parameters").OracleTextField("Errors Only").Enter RACK_GetData("Login_Data","Errors")
							OracleFlexWindow("Parameters").OracleTextField("Report").Enter RACK_GetData("Login_Data","Report")
							OracleFlexWindow("Parameters").OracleTextField("Transfer to General Ledger").Enter RACK_GetData("Login_Data","Transfer")
							OracleFlexWindow("Parameters").OracleTextField("Post in General Ledger").Enter RACK_GetData("Login_Data","Post")
							OracleFlexWindow("Parameters").OracleTextField("Include User Transaction").Enter RACK_GetData("Login_Data","IncludeUserTransaction")
												
						Case "Transfer journal entries to GL"							
							OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Login_Data","ReportName")
							OracleFlexWindow("Parameters").OracleTextField("Ledger").Enter RACK_GetData("Login_Data","Ledger") 							
							OracleFlexWindow("Parameters").OracleTextField("End Date").Enter RACK_GetData("Login_Data","EndDate")
							OracleFlexWindow("Parameters").OracleTextField("Post in General Ledger").Enter RACK_GetData("Login_Data","Post")
						Case "Invoice Aging report"		
							OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Login_Data","ReportName")
							OracleFlexWindow("Parameters").OracleTextField("Sort Invoices By").Enter RACK_GetData("Login_Data","SortBy")
							OracleFlexWindow("Parameters").OracleTextField("Include Invoice Detail").Enter RACK_GetData("Login_Data","IncludeInvoiceDetail")
							OracleFlexWindow("Parameters").OracleTextField("Include Site Detail").Enter RACK_GetData("Login_Data","IncludeSiteDetail")
							OracleFlexWindow("Parameters").OracleTextField("Aging Period Name").Enter RACK_GetData("Login_Data","Period") 					
						Case "Accounts Payable Trial Balance"		
							OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Login_Data","ReportName")
							OracleFlexWindow("Parameters").OracleTextField("Report Definition").Enter RACK_GetData("Login_Data","Report")
							OracleFlexWindow("Parameters").OracleTextField("Journal Source").Enter RACK_GetData("Login_Data","Source")
							OracleFlexWindow("Parameters").OracleTextField("As of Date").Enter RACK_GetData("Login_Data","AsOfDate")
							If OracleFlexWindow("Accounting Flexfield").OracleButton("Cancel").Exist(10) Then
								OracleFlexWindow("Accounting Flexfield").OracleButton("Cancel").Click
							End If
							OracleFlexWindow("Parameters").OracleTextField("Report Mode").SetFocus
							OracleFlexWindow("Parameters").OracleTextField("Report Mode").Enter RACK_GetData("Login_Data","Mode")	
						Case "Payables Posted Invoice Register"		
							OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Login_Data","ReportName")
							OracleFlexWindow("Parameters").OracleTextField("Ledger/Ledger Set").Enter RACK_GetData("Login_Data","Ledger") 
							OracleFlexWindow("Parameters").OracleTextField("Period From").Enter RACK_GetData("Login_Data","Period") 
							OracleFlexWindow("Parameters").OracleTextField("Order By").Enter RACK_GetData("Login_Data","OrderBy")	
							OracleFlexWindow("Parameters").OracleTextField("Include Manual entries").Enter RACK_GetData("Login_Data","IncludeManualentries")	
						Case "Payables Posted Payment Register"		
							OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Login_Data","ReportName")
							OracleFlexWindow("Parameters").OracleTextField("Ledger/Ledger Set").Enter RACK_GetData("Login_Data","Ledger") 
							OracleFlexWindow("Parameters").OracleTextField("Period From").Enter RACK_GetData("Login_Data","Period")  
							OracleFlexWindow("Parameters").OracleTextField("Order By").Enter RACK_GetData("Login_Data","OrderBy")	
						Case "Expense Report Export"		
							OracleFormWindow("Submit Request").OracleTextField("Run this Request...|Name").Enter RACK_GetData("Login_Data","ReportName")
							OracleFlexWindow("Parameters").OracleTextField("Source").Enter RACK_GetData("Login_Data","Source") 	
					End Select

					OracleFlexWindow("Parameters").Approve
					OracleFormWindow("Submit Request").OracleButton("Submit").Click	
					If OracleNotification("Decision").Exist(20) Then
						arrmessage=OracleNotification("Decision").GetROProperty("message")
						arrrequestID = Split (arrmessage,"=")
						requestID = Split(arrrequestID(1), ")")
						'MsgBox(Trim(requestID(0)))
						OracleNotification("Decision").Decline				
						OracleFormWindow("Navigator").SelectMenu "View->Requests"
						OracleFormWindow("Find Requests").OracleRadioGroup("My Completed Requests").Select "Specific Requests"
						
						OracleFormWindow("Find Requests").OracleTextField("Request ID").Enter Trim(requestID(0))
						OracleFormWindow("Find Requests").OracleButton("Find").Click
						counter =0
						Do
							wait(5)
							OracleFormWindow("Requests").OracleButton("Refresh Data").Click
							counter= counter+1				
						Loop Until (counter >RACK_GetData("Login_Data","ReportExeTime")) OR (OracleFormWindow("Requests").OracleTextField("Phase").GetROProperty("Value") = "Completed" AND OracleFormWindow("Requests").OracleTextField("Status").GetROProperty("Value") = "Normal" )
							If counter > RACK_GetData("Login_Data","ReportExeTime") Then
								 RACK_ReportEvent  "Report " ,"Request ID " &Trim(requestID(0)) & " Not completed . Check  the process manually", "Fail"
							Else
								OracleFormWindow("Requests").OracleButton("View Output").Click
								wait(2)
								RACK_ReportEvent   "Process" &Trim(requestID(0)) , "Sucessfully completed", "Pass"
								If Browser("Browser").Exist(5) Then
									wait(2)
										Browser("Browser").CloseAllTabs()
								End If
							End If
					Else 
							RACK_ReportEvent  "Report","Decision does not exist","Fail"
					End If
				Else
					RACK_ReportEvent  "Report","OK button does not exist","Fail"
				End If
			Else
				RACK_ReportEvent  "Report","Run link does not exist","Fail"
			End If
		Else
			RACK_ReportEvent  "Report","RS US Payables Clerk link does not exist","Fail"
		End If		
	End Select
End Sub
Sub Inventory_Receipt()
   If Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory Manager,").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory Manager,").Click
		If Browser("Login").Page("Oracle Applications Home_2").Link("Miscellaneous Transaction").Exist(20) Then
			Browser("Login").Page("Oracle Applications Home_2").Link("Miscellaneous Transaction").Click
			Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
			Browser("Login").Page("Oracle Applications Home_2").Sync
			If OracleListOfValues("Organizations").Exist(160) Then
				OracleListOfValues("Organizations").Select RACK_GetData("Login_Data", "Organization")
				If OracleFormWindow("Miscellaneous Transaction").OracleTextField("Transaction|Type").Exist(20) Then
					OracleFormWindow("Miscellaneous Transaction").OracleTextField("Transaction|Type").OpenDialog
					OracleListOfValues("Transaction Types").Select RACK_GetData("Login_Data", "Transaction")
					If OracleFormWindow("Miscellaneous Transaction").OracleTextField("Transaction|Source").Exist(20) Then
						OracleFormWindow("Miscellaneous Transaction").OracleTextField("Transaction|Source").OpenDialog
						OracleFormWindow("Miscellaneous Transaction").OracleCheckbox("Transaction|Serial-Triggered").Select
						OracleFormWindow("Miscellaneous Transaction").OracleButton("Transaction Lines").Click
						If OracleFormWindow("Account alias receipt").OracleTextField("Item").Exist(20) Then
							OracleFormWindow("Account alias receipt").OracleTextField("Item").Enter RACK_GetData("Login_Data", "item")
							OracleFormWindow("Account alias receipt").OracleTextField("Subinventory").Enter RACK_GetData("Login_Data", "Subinventory")
							OracleFormWindow("Account alias receipt").OracleTextField("Locator").Enter RACK_GetData("Login_Data", "Locator")
							OracleFormWindow("Account alias receipt").OracleTextField("Quantity").SetFocus
							OracleFormWindow("Account alias receipt").OracleTextField("Quantity").Enter RACK_GetData("Login_Data", "Qty")
							OracleFormWindow("Account alias receipt").OracleButton("Lot / Serial").Click
							If OracleFormWindow("Serial Entry").OracleTextField("Start Serial Number").Exist(20) Then
								OracleFormWindow("Serial Entry").OracleTextField("Start Serial Number").Enter RACK_GetData("Login_Data", "serialnumber")
								OracleFormWindow("Serial Entry").OracleButton("Done").Click
								OracleFormWindow("Account alias receipt").SelectMenu "File->Save"
							Else
								RACK_ReportEvent "Inventory Receipt", "Start Serial Number  text field does not exist","Fail"
							End If
						Else
							RACK_ReportEvent "Inventory Receipt", "Item  text field does not exist","Fail"
						End If
					Else
						RACK_ReportEvent "Inventory Receipt", "Transaction|Source  text field does not exist","Fail"
					End If
				Else
					RACK_ReportEvent "Inventory Receipt", "Transaction|Type text field does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Inventory Receipt", "Organizations List does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Inventory Receipt", "Miscellaneous Transaction Link does not exist","Fail"
		End If
	Else
		RACK_ReportEvent "Inventory Receipt", "RS US Inventory Manager Link does not exist","Fail"
   End If
End Sub
Sub Inventory_Issue()
   If Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory Manager,").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory Manager,").Click
		If Browser("Login").Page("Oracle Applications Home_2").Link("Miscellaneous Transaction").Exist(20) Then
			Browser("Login").Page("Oracle Applications Home_2").Link("Miscellaneous Transaction").Click
			Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
			Browser("Login").Page("Oracle Applications Home_2").Sync
			If OracleListOfValues("Organizations").Exist(160) Then
				OracleListOfValues("Organizations").Select RACK_GetData("Login_Data", "Organization")
				If OracleFormWindow("Miscellaneous Transaction").OracleTextField("Transaction|Type").Exist(20) Then
					OracleFormWindow("Miscellaneous Transaction").OracleTextField("Transaction|Type").OpenDialog			
					OracleListOfValues("Transaction Types").Select RACK_GetData("Login_Data", "Transaction")
					If OracleFormWindow("Miscellaneous Transaction").OracleTextField("Transaction|Source").Exist(20) Then
						OracleFormWindow("Miscellaneous Transaction").OracleTextField("Transaction|Source").OpenDialog
						OracleFormWindow("Miscellaneous Transaction").OracleCheckbox("Transaction|Serial-Triggered").Select
						OracleFormWindow("Miscellaneous Transaction").OracleButton("Transaction Lines").Click
						If OracleFormWindow("Account alias issue").OracleTextField("Item").Exist(20) Then
							OracleFormWindow("Account alias issue").OracleTextField("Item").Enter RACK_GetData("Login_Data", "item")
							OracleFormWindow("Account alias issue").OracleTextField("Subinventory").Enter RACK_GetData("Login_Data", "Subinventory")
							OracleFormWindow("Account alias issue").OracleTextField("Locator").Enter RACK_GetData("Login_Data", "Locator")
							OracleFormWindow("Account alias issue").OracleTextField("Quantity").SetFocus
							OracleFormWindow("Account alias issue").OracleTextField("Quantity").Enter RACK_GetData("Login_Data", "Qty")
							OracleFormWindow("Account alias issue").OracleButton("Lot / Serial").Click
							If OracleFormWindow("Serial Entry").OracleTextField("Start Serial Number").Exist(20) Then
								OracleFormWindow("Serial Entry").OracleTextField("Start Serial Number").Enter RACK_GetData("Login_Data", "serialnumber")
								OracleFormWindow("Serial Entry").OracleButton("Done").Click
								OracleFormWindow("Account alias issue").SelectMenu "File->Save"
							Else
								RACK_ReportEvent "Inventory Issue", "Start Serial Number  text field does not exist","Fail"
							End If
						Else
							RACK_ReportEvent "Inventory Issue", "Item  text field  does not exist","Fail"
						End If
					Else
						RACK_ReportEvent "Inventory Issue", "Transaction|Source  text field does not exist","Fail"
					End If
				Else
					RACK_ReportEvent "Inventory Issue", "Transaction|Type text field does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Inventory Issue", "Organizations List does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Inventory Issue", "Miscellaneous Transaction Link does not exist","Fail"
		End If
	Else
		RACK_ReportEvent "Inventory Issue", "RS US Inventory Manager Link does not exist","Fail"
   End If
End Sub
Sub BuildSubassembly()
If Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory User,").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory User,").Click
	If Browser("Login").Page("Oracle Applications Home_2").Link("Build Subassembly").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home_2").Link("Build Subassembly").Click
		If Browser("Login").Page("Configure Subassembly").WebList("InvOrganizationChoice").Exist(20) Then
			Browser("Login").Page("Configure Subassembly").WebList("InvOrganizationChoice").Click
			Browser("Login").Page("Configure Subassembly").WebList("InvOrganizationChoice").Select RACK_GetData("Login_Data", "Organization")
			If Browser("Login").Page("Configure Subassembly_2").WebEdit("ItemNo").Exist(20) Then
				Browser("Login").Page("Configure Subassembly_2").WebEdit("ItemNo").Set RACK_GetData("Login_Data", "item")
				Browser("Login").Page("Configure Subassembly_2").WebEdit("SerialNumber").Set RACK_GetData("Login_Data", "serialnumber")
				Browser("Login").Page("Configure Subassembly_2").WebButton("Next").Click
				If Browser("Login").Page("Configure Subassembly_3").WebEdit("PartsAdvTableRN:ChassisSerialNumber").Exist(20) Then
					Browser("Login").Page("Configure Subassembly_3").WebEdit("PartsAdvTableRN:ChassisSerialNumber").Set RACK_GetData("Login_Data", "ChassisSNo")
					Browser("Login").Page("Configure Subassembly_3").WebButton("Next").Click
					If Browser("Login").Page("Configure Subassembly_4").WebButton("Create").Exist(20) Then
						Browser("Login").Page("Configure Subassembly_4").WebButton("Create").Click
						status = Browser("Login").Page("Confirmation").WebElement("Subassembly Whitebox S3992").GetROProperty("innertext")
						Browser("Login").Page("Confirmation").WebButton("OK").Click
					Else
						RACK_ReportEvent "Build Subassembly", "Create button does not exist","Fail"
					End If
				Else
					RACK_ReportEvent "Build Subassembly", "PartsAdvTableRN:ChassisSerialNumber text box does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Build Subassembly", "ItemNo text box does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Build Subassembly", "InvOrganizationChoice list does not exist","Fail"
		End If
	Else
		RACK_ReportEvent "Build Subassembly", "Build Subassembly link does not exist","Fail"
	End If
Else
	RACK_ReportEvent "Build Subassembly", "RS US Inventory Use Link does not exist","Fail"
End If
End Sub
Sub Create_MasterItem()
   If Browser("Login").Page("Oracle Applications Home").Link("RS Inventory Item Master,").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home").Link("RS Inventory Item Master,").Click
		If Browser("Login").Page("Oracle Applications Home_2").Link("Master Items").Exist(20) Then
			Browser("Login").Page("Oracle Applications Home_2").Link("Master Items").Click
			Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
			Browser("Login").Page("Oracle Applications Home_2").Sync
			If OracleFormWindow("Master Item").Exist(100) Then
				OracleFormWindow("Master Item").SelectMenu "Tools->Copy From..."
				If OracleFormWindow("Copy From").OracleTextField("Template").Exist(20) Then
					OracleFormWindow("Copy From").OracleTextField("Template").Enter "RS Depreciable Item"
					OracleFormWindow("Copy From").OracleButton("Apply").Click
					OracleFormWindow("Copy From").OracleButton("Done").Click
					If OracleFormWindow("Master Item").OracleTextField("Item").Exist(20) Then
						OracleFormWindow("Master Item").OracleTextField("Item").Enter RACK_GetData("Login_Data", "item")
						OracleFormWindow("Master Item").OracleTextField("Description").Enter RACK_GetData("Login_Data", "Description")
						OracleFormWindow("Master Item").SelectMenu "File->Save"
						wait(1)
						Status = OracleStatusLine("OracleStatusLine").GetROProperty("message")
						'MsgBox(status)
					Else
						RACK_ReportEvent "Create Item", "Item text field does not exist","Fail"
					End If
				Else
					RACK_ReportEvent "Create Item", "Template text field does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Create Item", "Master Item form does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Create Item", "Master Items link does not exist","Fail"
		End If
	Else
		RACK_ReportEvent "Create Item", "RS Inventory Item Master link does not exist","Fail"
   End If
End Sub
Sub Update_MasterItem()
If Browser("Login").Page("Oracle Applications Home").Link("RS Inventory Item Master,").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS Inventory Item Master,").Click
	If Browser("Login").Page("Oracle Applications Home_2").Link("Master Items").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home_2").Link("Master Items").Click
		Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
		Browser("Login").Page("Oracle Applications Home_2").Sync
		If OracleFormWindow("Master Item").OracleTextField("Item").Exist(120) Then
			OracleFormWindow("Master Item").OracleTextField("Item").InvokeSoftkey "ENTER QUERY"
			OracleFormWindow("Master Item").OracleTextField("Item").Enter RACK_GetData("Login_Data", "item")
			OracleFormWindow("Master Item").OracleTextField("Item").InvokeSoftkey "EXECUTE QUERY"
			If OracleFormWindow("Master Item").OracleTabbedRegion("Main").OracleTextField("User Item Type").Exist(20) Then
				OracleFormWindow("Master Item").OracleTabbedRegion("Main").OracleTextField("User Item Type").Enter "RS Depreciable Item"
				OracleFormWindow("Master Item").SelectMenu "Tools->Categories"
				If OracleFormWindow("Category Assignment").OracleTextField("Category").Exist(20) Then
					OracleFormWindow("Category Assignment").OracleTextField("Category").OpenDialog
					OracleListOfValues("RS Item Categories").Select RACK_GetData("Login_Data", "Category")
					OracleFormWindow("Category Assignment").SelectMenu "File->Save"
				Else
					RACK_ReportEvent "Update Item", "Category text field does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Update Item", "User Item Type text field does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Update Item", "Item text field does not exist","Fail"
		End If
	Else
		RACK_ReportEvent "Update Item", "Master Items link does not exist","Fail"
	End If
Else
	RACK_ReportEvent "Update Item", "RS Inventory Item Master link does not exist","Fail"
End If
End Sub
Sub Inventory_OpenPeriod()
If Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory Accountant,").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory Accountant,").Click
	If Browser("Login").Page("Oracle Applications Home_2").Link("Inventory Accounting Periods").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home_2").Link("Inventory Accounting Periods").Click
		Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
		Browser("Login").Page("Oracle Applications Home_2").Sync
		If OracleListOfValues("Organizations").Exist(120) Then
			OracleListOfValues("Organizations").Select RACK_GetData("Login_Data", "Organization")
			OracleFormWindow("Inventory Accounting Periods").OracleButton("Change Status...").Click
			OracleNotification("Caution").Approve
		Else
			RACK_ReportEvent "Inventory OpenPeriod", "Organizations list does not exist","Fail"
		End If
	Else
		RACK_ReportEvent "Inventory OpenPeriod", "Inventory Accounting Periods link does not exist","Fail"
	End If
Else
	RACK_ReportEvent "Inventory OpenPeriod", "RS US Inventory Accountant  link does not exist","Fail"
End If
End Sub
Sub Inventory_ClosePeriod()
If Browser("Login").Page("Oracle Applications Home").Link("RS US Close Coordinator_3").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS US Close Coordinator_3").Click
	If Browser("Login").Page("Oracle Applications Home_2").Link("Inventory Accounting Periods").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home_2").Link("Inventory Accounting Periods").Click
		Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
		Browser("Login").Page("Oracle Applications Home_2").Sync
		If OracleListOfValues("Organizations").Exist(120) Then
			OracleListOfValues("Organizations").Select RACK_GetData("Login_Data", "Organization")
			If OracleFormWindow("Inventory Accounting Periods").OracleButton("Change Status...").Exist(20) Then
				OracleFormWindow("Inventory Accounting Periods").OracleButton("Change Status...").Click
				If OracleFormWindow("Change Period Status").OracleRadioGroup("New Status").Exist(10) Then
					OracleFormWindow("Change Period Status").OracleRadioGroup("New Status").Select "Closed  (Irreversible)"
					OracleFormWindow("Change Period Status").OracleButton("OK").Click
					If OracleNotification("Caution").Exist(20) Then
						OracleNotification("Caution").Approve
						OracleNotification("Note").Approve	
					Else
						RACK_ReportEvent "Inventory Close Period", "OK button does not exist","Fail"
					End If
				Else
					RACK_ReportEvent "Inventory Close Period", "New Status group does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Inventory Close Period", "Change Status button does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Inventory Close Period", "Organizations list does not exist","Fail"
		End If
	Else
		RACK_ReportEvent "Inventory Close Period", "Inventory Accounting Periods link does not exist","Fail"
	End If
Else
	RACK_ReportEvent "Inventory Close Period", "RS US Close Coordinator link does not exist","Fail"
End If
End Sub
Sub AP_OpenPeriod()
   If Browser("Login").Page("Oracle Applications Home").Link("RS US Close Coordinator_2").Exist(20) Then
        Browser("Login").Page("Oracle Applications Home").Link("RS US Close Coordinator_2").Click
		If Browser("Login").Page("Oracle Applications Home_2").Link("Control Payables Periods").Exist(20) Then
			Browser("Login").Page("Oracle Applications Home_2").Link("Control Payables Periods").Click
			Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
			Browser("Login").Page("Oracle Applications Home_2").Sync
			If OracleFormWindow("Control Payables Periods").OracleTextField("Period Status").Exist(120) Then
				OracleFormWindow("Control Payables Periods").OracleTextField("Period Status").InvokeSoftkey "ENTER QUERY"
				OracleFormWindow("Control Payables Periods").OracleTextField("Period Name").Enter RACK_GetData("Login_Data", "Period")
				OracleFormWindow("Control Payables Periods").OracleTextField("Period Name").InvokeSoftkey "EXECUTE QUERY"
				OracleFormWindow("Control Payables Periods").OracleTextField("Period Status").SetFocus
				OracleFormWindow("Control Payables Periods").OracleTextField("Period Status").OpenDialog
				If OracleListOfValues("Control Statuses").Exist(20) Then
					OracleListOfValues("Control Statuses").Select RACK_GetData("Login_Data", "Status")
					OracleFormWindow("Control Payables Periods").SelectMenu "File->Save" 
				Else
					RACK_ReportEvent "AP OpenPeriod", "Control Statuses list does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "AP OpenPeriod", "Period Status text field does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "AP OpenPeriod", "Control Payables Periods link does not exist","Fail"
		End If
	Else
		RACK_ReportEvent "AP OpenPeriod", "RS US Close Coordinator link does not exist","Fail"
   End If
End Sub
Sub AR_OpenPeriod()
If Browser("Login").Page("Oracle Applications Home").Link("RS US Close Coordinator").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS US Close Coordinator").Click
	If Browser("Login").Page("Oracle Applications Home_2").Link("Accounting Periods").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home_2").Link("Accounting Periods").Click
		Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
			Browser("Login").Page("Oracle Applications Home_2").Sync
		If OracleFormWindow("Find Receivables Periods").OracleTextField("Ledger").Exist(120) Then
			OracleFormWindow("Find Receivables Periods").OracleTextField("Ledger").Enter RACK_GetData("Login_Data", "Ledger")
        	OracleFormWindow("Find Receivables Periods").OracleButton("Find").Click
			If OracleFormWindow("Open/Close Accounting").OracleTextField("Status").Exist(20) Then
				OracleFormWindow("Open/Close Accounting").OracleTextField("Status").InvokeSoftkey "ENTER QUERY"
				OracleFormWindow("Open/Close Accounting").OracleTextField("Name").Enter RACK_GetData("Login_Data", "Period")
				OracleFormWindow("Open/Close Accounting").OracleTextField("Name").InvokeSoftkey "EXECUTE QUERY"
				If OracleFormWindow("Open/Close Accounting").OracleTextField("Status").Exist(20) Then
					OracleFormWindow("Open/Close Accounting").OracleTextField("Status").OpenDialog
					OracleListOfValues("Period Statuses").Select RACK_GetData("Login_Data", "Status")
					OracleFormWindow("Open/Close Accounting").SelectMenu "File->Save"
				Else
					RACK_ReportEvent "AR OpenPeriod", "Status text field does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "AR OpenPeriod", "Status text field does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "AR OpenPeriod", "Ledger text field does not exist","Fail"
		End If
	Else
		RACK_ReportEvent "AR OpenPeriod", "Accounting Periods link does not exist","Fail"
	End If
Else
	RACK_ReportEvent "AR OpenPeriod", "RS US Close Coordinator link does not exist","Fail"
End If
End Sub
Sub CreatePO()
Select Case RACK_GetData("Login_Data", "Transaction")
	Case "Blanket"
		If Browser("Login").Page("Oracle Applications Home").Link("RS US Purchasing Staff").Exist(20) Then
			Browser("Login").Page("Oracle Applications Home").Link("RS US Purchasing Staff").Click
			If Browser("Login").Page("Oracle Applications Home_2").Link("Purchase Orders").Exist(20) Then
				Browser("Login").Page("Oracle Applications Home_2").Link("Purchase Orders").Click
				Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
				Browser("Login").Page("Oracle Applications Home_2").Sync
				If OracleFormWindow("Purchase Orders").OracleTextField("Type").Exist(120) Then
					OracleFormWindow("Purchase Orders").OracleTextField("Type").Enter "Blanket Purchase Agreement"
					If OracleFormWindow("Purchase Orders").OracleTextField("Supplier").Exist(20) Then
						OracleFormWindow("Purchase Orders").OracleTextField("Supplier").SetFocus
						OracleFormWindow("Purchase Orders").OracleTextField("Supplier").Enter RACK_GetData("Login_Data", "Supplier")
						OracleFormWindow("Purchase Orders").OracleTextField("Site").SetFocus
						OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Item").SetFocus
						OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Item").Enter RACK_GetData("Login_Data", "item")
						OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Quantity").SetFocus
						OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Quantity").Enter RACK_GetData("Login_Data", "Qty")
						OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Price").Enter RACK_GetData("Login_Data", "Price")
						OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Need-By").SetFocus
						OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Need-By").Enter RACK_GetData("Login_Data", "NeedBy")
						OracleFormWindow("Purchase Orders").SelectMenu "File->Save"
						PO_ID = OracleFormWindow("Purchase Orders").OracleTextField("PO, Rev").GetROProperty("value")
						'MsgBox(PO_ID)
						Environment.Value("POID") = Trim(PO_ID)
						RACK_ReportEvent "Create PO", "PO "&PO_ID&" Created succecssfully","Pass"
						If OracleFormWindow("Purchase Orders").OracleButton("Approve...").Exist(20) Then
							OracleFormWindow("Purchase Orders").OracleButton("Approve...").Click
							OracleFormWindow("Approve Document").OracleTabbedRegion("Approval Details").OracleCheckbox("Approval|Submit for Approval").Select
							OracleFormWindow("Approve Document").OracleButton("OK").Click
							 status = OracleFormWindow("Purchase Orders").OracleTextField("Status").GetROProperty("value")
							RACK_ReportEvent  " PO status", "PO - " & status, "Pass"
						Else	
							RACK_ReportEvent "Create PO", "Approve button does not exist","Fail"
						End If						
					Else
						RACK_ReportEvent "Create PO", "Supplier text field does not exist","Fail"
					End If
				Else
					RACK_ReportEvent "Create PO", "Supplier text field does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Create PO", "Purchase Orders link does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Create PO", "RS US Purchasing Staff link does not exist","Fail"
		End If
	Case "Notification"		
		If Browser("Login").Page("Oracle Applications Home").Link("RS US Purchasing Staff").Exist(20) Then
			Browser("Login").Page("Oracle Applications Home").Link("RS US Purchasing Staff").Click
			If Browser("Login").Page("Oracle Applications Home_2").Link("Purchase Orders").Exist(20) Then
				Browser("Login").Page("Oracle Applications Home_2").Link("Purchase Orders").Click
				Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
				Browser("Login").Page("Oracle Applications Home_2").Sync
				If OracleFormWindow("Purchase Orders").OracleTextField("Supplier").Exist(120) Then
					OracleFormWindow("Purchase Orders").OracleTextField("Supplier").SetFocus
					OracleFormWindow("Purchase Orders").OracleTextField("Supplier").Enter RACK_GetData("Login_Data", "Supplier")
					OracleFormWindow("Purchase Orders").OracleTextField("Site").SetFocus
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Type").SetFocus
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Type").Enter "Goods"
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Item").SetFocus
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Item").Enter RACK_GetData("Login_Data", "item")
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Quantity").SetFocus
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Quantity").Enter RACK_GetData("Login_Data", "Qty")
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Price").Enter RACK_GetData("Login_Data", "Price")
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Need-By").SetFocus
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Need-By").Enter RACK_GetData("Login_Data", "NeedBy")
					OracleFormWindow("Purchase Orders").SelectMenu "File->Save"										
					PO_ID = OracleFormWindow("Purchase Orders").OracleTextField("PO, Rev").GetROProperty("value")
					'MsgBox(PO_ID)
					Environment.Value("POID") = Trim(PO_ID)
					RACK_ReportEvent "Create PO", "PO "&PO_ID&" Created succecssfully","Pass"
					If OracleFormWindow("Purchase Orders").OracleButton("Approve...").Exist(20) Then
						OracleFormWindow("Purchase Orders").OracleButton("Approve...").Click
						OracleFormWindow("Approve Document").OracleTabbedRegion("Approval Details").OracleCheckbox("Approval|Submit for Approval").Select
						OracleFormWindow("Approve Document").OracleTabbedRegion("Approval Details").OracleCheckbox("Approval|Forward").Select
						OracleFormWindow("Approve Document").OracleTabbedRegion("Approval Details").OracleTextField("Approval|Forward To").Enter RACK_GetData("Login_Data", "ForwardTo")
						OracleFormWindow("Approve Document").OracleButton("OK").Click
					 Else	
						RACK_ReportEvent "Create PO", "Approve button does not exist","Fail"
					End If
				Else
					RACK_ReportEvent "Create PO", "Supplier text field does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Create PO", "Purchase Orders link does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Create PO", "RS US Purchasing Staff link does not exist","Fail"
		End If
	Case Else
		If Browser("Login").Page("Oracle Applications Home").Link("RS US Purchasing Staff").Exist(20) Then
			Browser("Login").Page("Oracle Applications Home").Link("RS US Purchasing Staff").Click
			If Browser("Login").Page("Oracle Applications Home_2").Link("Purchase Orders").Exist(20) Then
				Browser("Login").Page("Oracle Applications Home_2").Link("Purchase Orders").Click
				Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
				Browser("Login").Page("Oracle Applications Home_2").Sync
				If OracleFormWindow("Purchase Orders").OracleTextField("Supplier").Exist(120) Then
					OracleFormWindow("Purchase Orders").OracleTextField("Supplier").SetFocus
					OracleFormWindow("Purchase Orders").OracleTextField("Supplier").Enter RACK_GetData("Login_Data", "Supplier")
					OracleFormWindow("Purchase Orders").OracleTextField("Site").SetFocus
					If RACK_GetData("Login_Data", "Type" ) = "Expense" Then
						OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Type").SetFocus
						OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Type").Enter "Service"
					Else
						OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Type").SetFocus
						OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Type").Enter "Goods"
					End If
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Item").SetFocus
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Item").Enter RACK_GetData("Login_Data", "item")
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Quantity").SetFocus
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Quantity").Enter RACK_GetData("Login_Data", "Qty")
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Price").Enter RACK_GetData("Login_Data", "Price")
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Need-By").SetFocus
					OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Need-By").Enter RACK_GetData("Login_Data", "NeedBy")
					If OracleFormWindow("Purchase Orders").OracleButton("Shipments").Exist(10) Then
						OracleFormWindow("Purchase Orders").OracleButton("Shipments").Click
						If OracleFormWindow("Shipments").OracleButton("Distributions").Exist(50) Then
							OracleFormWindow("Shipments").OracleButton("Distributions").Click	
							OracleFormWindow("Purchase Orders").SelectMenu "File->Save"										
							PO_ID = OracleFormWindow("Purchase Orders").OracleTextField("PO, Rev").GetROProperty("value")
							'MsgBox(PO_ID)
							Environment.Value("POID") = Trim(PO_ID)
							RACK_ReportEvent "Create PO", "PO "&PO_ID&" Created succecssfully","Pass"
							If OracleFormWindow("Purchase Orders").OracleButton("Approve...").Exist(20) Then
								OracleFormWindow("Purchase Orders").OracleButton("Approve...").Click
								OracleFormWindow("Approve Document").OracleTabbedRegion("Approval Details").OracleCheckbox("Approval|Submit for Approval").Select
								OracleFormWindow("Approve Document").OracleButton("OK").Click
								 status = OracleFormWindow("Purchase Orders").OracleTextField("Status").GetROProperty("value")
								RACK_ReportEvent  " PO status", "PO - " & status, "Pass"
							Else	
								RACK_ReportEvent "Create PO", "Approve button does not exist","Fail"
							End If
						End If
					End If
				Else
					RACK_ReportEvent "Create PO", "Supplier text field does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Create PO", "Purchase Orders link does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Create PO", "RS US Purchasing Staff link does not exist","Fail"
		End If
	End Select
End Sub
Sub ForwardDocument()
   	Call CreatePO()
	If OracleFormWindow("Purchase Orders").Exist(20) Then
		OracleFormWindow("Purchase Orders").CloseWindow
		OracleFormWindow("Navigator").SelectMenu "File->Switch Responsibility..."
		OracleListOfValues("Responsibilities").Select "RS US Purchasing Supervisor"
		OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List").Select "+  Management"
		OracleFormWindow("Navigator").OracleButton("Open").Click
		If OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List").Exist(20) Then
			OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List").Select "       Forward Documents"
			OracleFormWindow("Navigator").OracleButton("Open").Click
			If OracleFormWindow("Find Documents").OracleTextField("Document Number").Exist(20) Then
				OracleFormWindow("Find Documents").OracleTextField("Document Number").Enter Environment.Value("POID")
				wait(2)
				If OracleFormWindow("Find Documents").OracleButton("OK").Exist(20) Then
					OracleFormWindow("Find Documents").OracleButton("OK").Click
					If OracleNotification("Forms").OracleButton("OK").Exist(20) Then
						OracleNotification("Forms").OracleButton("OK").Click
					End If
					If OracleFormWindow("Forward Documents").OracleTextField("New Approver").Exist(20) Then
						OracleFormWindow("Forward Documents").OracleTextField("New Approver").Enter RACK_GetData("Login_Data", "Name")
						OracleFormWindow("Forward Documents").OracleCheckbox("OracleCheckbox").Select
						OracleFormWindow("Forward Documents").SelectMenu "File->Save"
						If OracleNotification("Forms").OracleButton("OK").Exist(20) Then
							OracleNotification("Forms").OracleButton("OK").Click	
							RACK_ReportEvent "Forward  Documents", "Document has been forwarded to " &  RACK_GetData("Login_Data", "Name") ,"Pass"
						Else
							RACK_ReportEvent "Forward  Documents", "Document has not been forwarded ","Fail"
						End If		
					Else	
						RACK_ReportEvent "Forward  Documents", "New Approver text field does not exist","Fail"
					End If
				Else
					RACK_ReportEvent "Forward  Documents", "OK button does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Forward  Documents", "Document Number text field does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Forward  Documents", "Function List does not exist","Fail"
		End If
	Else
		RACK_ReportEvent "Forward  Documents", "Purchase Orders form does not exist","Fail"
	End If
End Sub
Sub VerifyDocuments()
   If OracleFormWindow("Forward Documents").Exist(20) Then
		OracleFormWindow("Forward Documents").CloseWindow
		OracleFormWindow("Navigator").SelectMenu "File->Switch Responsibility..."
		OracleListOfValues("Responsibilities").Select "RS US Purchasing Staff"
		OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List").Select "+  Purchase Orders"
		OracleFormWindow("Navigator").OracleButton("Open").Click
		If OracleFormWindow("Navigator").OracleButton("Open").Exist(20) Then
			OracleFormWindow("Navigator").OracleButton("Open").Click
			If OracleFormWindow("Find Purchase Orders").OracleTextField("Number").Exist(20) Then
				OracleFormWindow("Find Purchase Orders").OracleTextField("Number").Enter Environment.Value("POID")
				OracleFormWindow("Find Purchase Orders").OracleButton("Find").Click
				OracleFormWindow("Purchase Order Headers").SelectMenu "Inquire->View Action History"
				namevalue = OracleFormWindow("Standard Purchase Order").OracleTextField("Performed By").GetROProperty("value")
                If StrComp( RACK_GetData("Login_Data", "ForwardTo"), namevalue,1) = 0 Then
					RACK_ReportEvent "Verify  Documents", "Document has been forwarded ","Pass"
				Else
					RACK_ReportEvent "Verify  Documents", "Document has not been forwarded ","Fail"
				End If
			Else
				RACK_ReportEvent "Verify  Documents", "Number text field does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Verify  Documents", "Open button does not exist","Fail"
		End If
	Else
		RACK_ReportEvent "Verify  Documents", "Forward Documents form does not exist","Fail"
   End If
End Sub
'Sub Create_item()
'	If Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory Administrator,").Exist(20) Then
'		Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory Administrator,").Click
'		If Browser("Login").Page("Oracle Applications Home_2").Link("Master Items").Exist(20) Then
'			Browser("Login").Page("Oracle Applications Home_2").Link("Master Items").Click
'			Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
'			Browser("Login").Page("Oracle Applications Home_2").Sync
'			If OracleListOfValues("Organizations").Exist(120) Then
'				OracleListOfValues("Organizations").Select RACK_GetData("Login_Data", "Organization")
'				OracleFormWindow("Master Item").OracleTextField("Item").Enter RACK_GetData("Login_Data", "item")	
'				OracleFormWindow("Master Item").OracleTextField("Description").Enter RACK_GetData("Login_Data", "Description")
'				OracleFormWindow("Master Item").SelectMenu "File->Save"	
'			Else
'				RACK_ReportEvent "Create item", "Organizations list does not exist","Fail"
'			End If
'		Else
'			RACK_ReportEvent "Create item", "Master Items link does not exist","Fail"
'		End If
'	Else
'		RACK_ReportEvent "Create item", "RS US Inventory Administrator link does not exist","Fail"
'	End If
'End Sub
Sub AssignOrganisation()
	'Call Create_item()
	If OracleFormWindow("Master Item").Exist(20) Then
		OracleFormWindow("Master Item").SelectMenu "Tools->Organization Assignment"
		If OracleFormWindow("Master Item").OracleButton("Assign All").Exist(20) Then
			OracleFormWindow("Master Item").OracleButton("Assign All").Click
			OracleFormWindow("Master Item").SelectMenu "File->Save"
		Else
			RACK_ReportEvent "Assign Organisation", "Assign All button does not exist","Fail"
		End If
	Else
		RACK_ReportEvent "Assign Organisation", "Master Item form does not exist","Fail"
	End If
End Sub
Sub InventoryReceipts()
	If RACK_GetData("Login_Data", "Transaction") = "PO" Then
		Call CreatePO()
		OracleFormWindow("Purchase Orders").CloseForm
		OracleFormWindow("Navigator").SelectMenu "File->Switch Responsibility..."
		OracleListOfValues("Responsibilities").Select "RS US Inventory User"
		OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleButton("Expand Branch").Click
		OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List").Select "          Receipts"
		OracleFormWindow("Navigator").OracleButton("Open").Click
		If OracleListOfValues("Organizations").Exist(120) Then
			OracleListOfValues("Organizations").Select  RACK_GetData("Login_Data", "Organization")
			If OracleFormWindow("Find Expected Receipts").OracleButton("Clear").Exist(20) Then
				OracleFormWindow("Find Expected Receipts").OracleButton("Clear").Click
				OracleFormWindow("Find Expected Receipts").OracleTabbedRegion("Supplier and Internal").OracleTextField("Purchase Order").Enter Environment.Value("POID")
				OracleFormWindow("Find Expected Receipts").OracleButton("Find").Click
				If OracleFormWindow("Receipt Header").OracleTextField("Receipt Date").Exist(20) Then					
					OracleFormWindow("Receipt Header").CloseWindow
					If RACK_GetData("Login_Data", "Type" ) = "Expense" Then
						OracleFormWindow("Receipts").OracleTabbedRegion("Lines").OracleTextField("Destination Type").OpenDialog
						OracleListOfValues("Destination Types").Select "Receiving"
					Else
						OracleFormWindow("Receipts").OracleTabbedRegion("Lines").OracleTextField("Destination Type").OpenDialog
						OracleListOfValues("Destination Types").Select "Inventory"
						OracleFormWindow("Receipts").OracleTabbedRegion("Lines").OracleTextField("Location").Enter "SA3-Datapoint Drive"
						OracleFormWindow("Receipts").OracleTabbedRegion("Lines").OracleTextField("Subinventory").Enter RACK_GetData("Login_Data", "Subinventory")
						OracleFormWindow("Receipts").OracleTabbedRegion("Lines").OracleTextField("Locator").Enter RACK_GetData("Login_Data", "Locator")
						If OracleFormWindow("Receipts").OracleButton("Lot - Serial").Exist(20) Then
							OracleFormWindow("Receipts").OracleButton("Lot - Serial").Click
							OracleFormWindow("Serial Entry").OracleTextField("Start Serial Number").Enter RACK_GetData("Login_Data", "serialnumber")
							OracleFormWindow("Serial Entry").OracleButton("Done").Click
						Else
							RACK_ReportEvent "Inventory Receipts", "Lot - Serial button does not exist","Fail"
						End If
					End If
						OracleFormWindow("Receipts").SelectMenu "File->Save"
						If OracleFormWindow("Receipts").OracleButton("Header").Exist(20) Then
							OracleFormWindow("Receipts").OracleButton("Header").Click
							ReceiptID = OracleFormWindow("Receipt Header").OracleTextField("Receipt").GetROProperty("value")
							'MsgBox(ReceiptID)
							RACK_ReportEvent "Inventory Receipts", "Receipt  "&ReceiptID&" Created succecssfully","Pass"
						Else
							RACK_ReportEvent "Inventory Receipts", "Header button does not exist","Fail"
						End If						
				Else
					RACK_ReportEvent "Inventory Receipts", "Receipt Date text field does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Inventory Receipts", "Clear button does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Inventory Receipts", "Organizations list does not exist","Fail"
		End If
	Else
		Call InterOrgShippment()
		OracleFormWindow("Intransit Shipment").CloseForm
		OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleButton("Expand Branch").Click
		OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List").Select "          Receipts"
		OracleFormWindow("Navigator").OracleButton("Open").Click
		If OracleListOfValues("Organizations").Exist(120) Then
			OracleListOfValues("Organizations").Select  RACK_GetData("Login_Data", "Organization")
			OracleFormWindow("Find Expected Receipts").OracleTabbedRegion("Supplier and Internal").OracleTextField("Shipment").Enter  RACK_GetData("Login_Data", "ShippmentNo")
			OracleFormWindow("Find Expected Receipts").OracleButton("Find").Click
			If OracleFormWindow("Receipt Header").OracleTextField("Receipt Date").Exist(20) Then
				'OracleFormWindow("Receipt Header").OracleTextField("Receipt Date").Enter "30-SEP-2012"
				OracleFormWindow("Receipt Header").CloseWindow
				OracleFormWindow("Receipts").OracleTabbedRegion("Lines").OracleTextField("Destination Type").Enter "Inventory"      					
				OracleFormWindow("Receipts").OracleTabbedRegion("Lines").OracleTextField("Subinventory").Enter RACK_GetData("Login_Data", "Subinventory")
				OracleFormWindow("Receipts").OracleTabbedRegion("Lines").OracleTextField("Locator").Enter RACK_GetData("Login_Data", "Locator")
				If OracleFormWindow("Receipts").OracleButton("Lot - Serial").Exist(20) Then
					OracleFormWindow("Receipts").OracleButton("Lot - Serial").Click
					OracleFormWindow("Serial Entry").OracleTextField("Start Serial Number").Enter RACK_GetData("Login_Data", "serialnumber")
					OracleFormWindow("Serial Entry").OracleButton("Done").Click
				Else
					RACK_ReportEvent "Inventory Receipts", "Lot - Serial button does not exist","Fail"
				End If			
				OracleFormWindow("Receipts").SelectMenu "File->Save"
				If OracleFormWindow("Receipts").OracleButton("Header").Exist(20) Then
					OracleFormWindow("Receipts").OracleButton("Header").Click
					ReceiptID = OracleFormWindow("Receipt Header").OracleTextField("Receipt").GetROProperty("value")
					'MsgBox(ReceiptID)
					RACK_ReportEvent "Inventory Receipts", "Receipt  "&ReceiptID&" Created succecssfully","Pass"
				Else
					RACK_ReportEvent "Inventory Receipts", "Header button does not exist","Fail"
				End If						
			Else
				RACK_ReportEvent "Inventory Receipts", "Receipt Date text field does not exist","Fail"
			End If			
		Else
			RACK_ReportEvent "Inventory Receipts", "Organizations list does not exist","Fail"
		End If
	End If
End Sub

Function Invoice()
Invoice = CStr(Int((999999-100000+1)*Rnd+100000))
End Function
Sub CreateInvoice()
   Select Case RACK_GetData("Login_Data", "Type" )
		Case "NonPO"
			If Browser("Login").Page("Oracle Applications Home").Link("RS US Payables Clerk").Exist(20) Then
				Browser("Login").Page("Oracle Applications Home").Link("RS US Payables Clerk").Click
				RACK_ReportEvent "Create Invoice", "RS US Payables Clerk link is clicked","Pass"
				If Browser("Login").Page("Oracle Applications Home_2").Link("Invoice Batches").Exist(20) Then
					Browser("Login").Page("Oracle Applications Home_2").Link("Invoice Batches").Click
					RACK_ReportEvent "Create Invoice", "Invoice Batches link is clicked","Pass"
					Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
					Browser("Login").Page("Oracle Applications Home_2").Sync
					If OracleFormWindow("Invoice Batches").OracleTextField("Batch Name").Exist(120) Then
						OracleFormWindow("Invoice Batches").OracleTextField("Batch Name").Enter RACK_GetData("Login_Data", "BatchName")
						RACK_ReportEvent "Create Invoice", RACK_GetData("Login_Data", "BatchName") & " is entered as batch name","Pass"
						OracleFormWindow("Invoice Batches").OracleButton("Invoices").Click
						If OracleFormWindow("Invoice Workbench").OracleTextField("Trading Partner").Exist(20) Then
							OracleFormWindow("Invoice Workbench").SelectMenu "Folder->Show Field..."
							OracleListOfValues("Show Field").Find "Requester"
							OracleListOfValues("Show Field").Select "Requester                                                                                                                                                                                                                  *REQUESTER_NAME"
							OracleFormWindow("Invoice Workbench").OracleTextField("Requester").Enter RACK_GetData("Login_Data", "Name")
							OracleFormWindow("Invoice Workbench").OracleTextField("Trading Partner").Enter RACK_GetData("Login_Data", "Supplier")
							RACK_ReportEvent "Create Invoice", RACK_GetData("Login_Data", "Supplier") &" is entered as supplier","Pass"
							OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Date").SetFocus
							OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Num").SetFocus
							Environment.Value("InvoiceID") = Invoice()
							OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Num").Enter Environment.Value("InvoiceID")
							RACK_ReportEvent "Create Invoice", Environment.Value("InvoiceID")& " is entered as invoice number" ,"Pass"
							OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Amount").Enter RACK_GetData("Login_Data", "Price")
							RACK_ReportEvent "Create Invoice", RACK_GetData("Login_Data", "Price")& " is entered as amount","Pass"
							OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Amount").SetFocus
							OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Amount").Enter RACK_GetData("Login_Data", "Price")
							If RACK_GetData("Login_Data", "Tax") = "Tax Accrual" Then
								OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Tax Classification Code").SetFocus
								OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Tax Classification Code").Enter RACK_GetData("Login_Data", "TaxCode")
							End If
							OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleButton("Distributions").Click
							RACK_ReportEvent "Create Invoice",  "Distributions buuton is clicked","Pass"
							If OracleFormWindow("Distributions").OracleTextField("Amount").Exist(20) Then
								OracleFormWindow("Distributions").OracleTextField("Amount").Enter RACK_GetData("Login_Data", "Price")
								OracleFormWindow("Distributions").OracleTextField("Account").SetFocus
								OracleFormWindow("Distributions").OracleTextField("Account").Enter RACK_GetData("Login_Data", "Account")
								RACK_ReportEvent "Create Invoice",RACK_GetData("Login_Data", "Account")& " is entered as account","Pass"
								OracleFormWindow("Distributions").SelectMenu "File->Save"
								OracleFormWindow("Distributions").CloseWindow
								If RACK_GetData("Login_Data", "Freight") = "Yes" Then
									OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Num").SetFocus
									OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Type").SetFocus
									OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Type").OpenDialog
									OracleListOfValues("Line Types").Select "Freight"
									RACK_ReportEvent "Create Invoice", "Freight details are entered","Pass"
									OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Amount_2").Enter RACK_GetData("Login_Data", "Price")
									RACK_ReportEvent "Create Invoice",RACK_GetData("Login_Data", "Price")& " is enter as freight amount","Pass"
									OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleButton("Allocations").Click
									OracleFormWindow("Allocation Rules").OracleButton("OK").Click
								End If
								If RACK_GetData("Login_Data", "Miscellaneous") = "Yes" Then
									OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Num_2").SetFocus
									OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Type_2").SetFocus
									OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Type_2").OpenDialog
									OracleListOfValues("Line Types").Select "Miscellaneous"
									RACK_ReportEvent "Create Invoice", "Miscellaneous details are entered","Pass"
									OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Amount_3").Enter RACK_GetData("Login_Data", "Price")
									RACK_ReportEvent "Create Invoice", RACK_GetData("Login_Data", "Price") & " is entered as mscellaneous ","Pass"
									OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleButton("Allocations").Click
									OracleFormWindow("Allocation Rules").OracleButton("OK").Click
								End If
								If RACK_GetData("Login_Data", "Tax") = "Yes" Then
									OracleFormWindow("Invoice Workbench").OracleButton("Tax Details").Click
									OracleFormWindow("Detail Tax Lines").OracleTextField("Tax  Regime  Code").SetFocus
									OracleFormWindow("Detail Tax Lines").OracleTextField("Tax  Regime  Code").OpenDialog
									OracleListOfValues("Tax Regimes").Select RACK_GetData("Login_Data", "TaxRegimes")
									RACK_ReportEvent "Create Invoice",RACK_GetData("Login_Data", "TaxRegimes")& " is entered","Pass"
									OracleFormWindow("Detail Tax Lines").OracleTextField("Tax").OpenDialog
									OracleFormWindow("Detail Tax Lines").OracleTextField("Tax Status").OpenDialog
									OracleFormWindow("Detail Tax Lines").OracleTextField("Rate Name").OpenDialog
									OracleListOfValues("Tax Rates").Select RACK_GetData("Login_Data", "TaxRates")
									RACK_ReportEvent "Create Invoice", RACK_GetData("Login_Data", "TaxRates")&" is entered","Pass"
									OracleFormWindow("Detail Tax Lines").OracleButton("OK").Click
									OracleFormWindow("Detail Tax Lines").OracleTextField("Tax Amount").SetFocus
									OracleNotification("Forms").Approve
									OracleFormWindow("Detail Tax Lines").OracleTextField("Place of Supply").Enter "BILL_FROM"
									OracleFormWindow("Detail Tax Lines").OracleButton("OK").Click									
								End If
								
								totalamt = OracleFormWindow("Invoice Workbench").OracleTabbedRegion("1 General").OracleTextField("Summary|Total").GetROProperty("value")
								OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Amount").SetFocus
								OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Amount").Enter totalamt
								OracleFormWindow("Invoice Workbench").SelectMenu "File->Save"
								RACK_ReportEvent "Create Invoice", "Invoice has been created successfully " &Environment.Value("InvoiceID"),"Pass"
								RACK_ReportEvent "Create Invoice", "Invoice has been created successfully ","Screenshot"
		'						OracleFormWindow("Invoice Workbench").SelectMenu "View->Attachments..."
		'						If OracleFormWindow("Attachments").OracleTabbedRegion("Source").OracleTextField("Category").Exist(20) Then
		'							OracleFormWindow("Attachments").OracleTabbedRegion("Source").OracleTextField("Category").Enter "misc"
		'							OracleFormWindow("Attachments").OracleTabbedRegion("Source").OracleTextField("Data Type").Enter "file"
		'							If Browser("GFM Upload Page").Page("GFM Upload Page").WebFile("Upload_oafileUpload").Exist(20) Then
		'								Browser("GFM Upload Page").Page("GFM Upload Page").WebFile("Upload_oafileUpload").Set "C:\Users\221045\Desktop\sample.txt"
		'								Browser("GFM Upload Page").Page("GFM Upload Page").WebButton("Submit").Click
		'								If Browser("GFM Upload Page").Page("GFM Upload").Link("Close Window").Exist(20) Then
		'									Browser("GFM Upload Page").Page("GFM Upload").Link("Close Window").Click
		'									Browser("Login").Page("Oracle Applications Home_2").Sync
		'									If OracleNotification("Decision").Exist(20) Then
		'										OracleNotification("Decision").Approve
		'										OracleFormWindow("Attachments").CloseWindow                 											
		'									Else
		'										RACK_ReportEvent "Create Invoice", "Decision  does not exist","Fail"
		'									End If
		'								Else
		'									RACK_ReportEvent "Create Invoice", "Close window link does not exist","Fail"
		'								End If
		'							Else
		'								RACK_ReportEvent "Create Invoice", "Upload_oafileUpload file does not exist","Fail"
		'							End If
		'						Else
		'							RACK_ReportEvent "Create Invoice", "Category text field does not exist","Fail"
		'						End If
							Else
								RACK_ReportEvent "Create Invoice", "Amount  text field does not exist","Fail"
							End If
						Else
							RACK_ReportEvent "Create Invoice", "Trading Partner text field does not exist","Fail"
						End If
					Else
						RACK_ReportEvent "Create Invoice", "Batch Name  text field does not exist","Fail"
					End If
				Else
					RACK_ReportEvent "Create Invoice", "Invoice batches link does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Create Invoice", "RS US Payables Clerk link does not exist","Fail"
		   End If
		Case "Credit Memo"
			 If Browser("Login").Page("Oracle Applications Home").Link("RS US Payables Clerk").Exist(20) Then
				Browser("Login").Page("Oracle Applications Home").Link("RS US Payables Clerk").Click
				RACK_ReportEvent "Create Invoice", "RS US Payables Clerk link is clicked","Pass"
				If Browser("Login").Page("Oracle Applications Home_2").Link("Invoice Batches").Exist(20) Then
					Browser("Login").Page("Oracle Applications Home_2").Link("Invoice Batches").Click
					RACK_ReportEvent "Create Invoice", "Invoice Batches link is clicked","Pass"
					Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
					Browser("Login").Page("Oracle Applications Home_2").Sync
					If OracleFormWindow("Invoice Batches").OracleTextField("Batch Name").Exist(120) Then
						OracleFormWindow("Invoice Batches").OracleTextField("Batch Name").Enter RACK_GetData("Login_Data", "BatchName")
						RACK_ReportEvent "Create Invoice", RACK_GetData("Login_Data", "BatchName") & " is entered as batch name","Pass"
						OracleFormWindow("Invoice Batches").OracleButton("Invoices").Click
						OracleFormWindow("Invoice Workbench").OracleTextField("Type").Enter RACK_GetData("Login_Data", "Type" )			
						OracleNotification("Note").Approve
						If OracleFormWindow("Invoice Workbench").OracleTextField("Trading Partner").Exist(20) Then
'							OracleFormWindow("Invoice Workbench").SelectMenu "Folder->Show Field..."
'							OracleListOfValues("Show Field").Find "Requester"
'							OracleListOfValues("Show Field").Select "Requester                                                                                                                                                                                                                  *REQUESTER_NAME"
'							OracleFormWindow("Invoice Workbench").OracleTextField("Requester").Enter RACK_GetData("Login_Data", "Name")
							OracleFormWindow("Invoice Workbench").OracleTextField("Trading Partner").Enter RACK_GetData("Login_Data", "Supplier")
							RACK_ReportEvent "Create Invoice", RACK_GetData("Login_Data", "Supplier") &" is entered as supplier","Pass"
							OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Date").SetFocus
							OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Num").SetFocus
							Environment.Value("InvoiceID") = Invoice()
							OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Num").Enter Environment.Value("InvoiceID")
							RACK_ReportEvent "Create Invoice", Environment.Value("InvoiceID")& " is entered as invoice number" ,"Pass"
							OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Amount").Enter RACK_GetData("Login_Data", "Price")
							RACK_ReportEvent "Create Invoice", RACK_GetData("Login_Data", "Price")& " is entered as amount","Pass"
							OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Amount").SetFocus
							OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Amount").Enter RACK_GetData("Login_Data", "Price")
							OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleButton("Distributions").Click
							RACK_ReportEvent "Create Invoice",  "Distributions buuton is clicked","Pass"
							If OracleFormWindow("Distributions").OracleTextField("Amount").Exist(20) Then
								OracleFormWindow("Distributions").OracleTextField("Amount").Enter RACK_GetData("Login_Data", "Price")
								OracleFormWindow("Distributions").OracleTextField("Account").SetFocus
								OracleFormWindow("Distributions").OracleTextField("Account").Enter RACK_GetData("Login_Data", "Account")
								RACK_ReportEvent "Create Invoice",RACK_GetData("Login_Data", "Account")& " is entered as account","Pass"
								OracleFormWindow("Distributions").SelectMenu "File->Save"
								OracleFormWindow("Distributions").CloseWindow						
								RACK_ReportEvent "Create Invoice", "Credit Memo has been created successfully. Invoice number is  " &Environment.Value("InvoiceID"),"Pass"		
								RACK_ReportEvent "Create Invoice", "Credit Memo has been created successfully ","Screenshot"
							Else
								RACK_ReportEvent "Create Invoice", "Amount  text field does not exist","Fail"
							End If
						Else
							RACK_ReportEvent "Create Invoice", "Trading Partner text field does not exist","Fail"
						End If
					Else
						RACK_ReportEvent "Create Invoice", "Batch Name  text field does not exist","Fail"
					End If
				Else
					RACK_ReportEvent "Create Invoice", "Invoice batches link does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Create Invoice", "RS US Payables Clerk link does not exist","Fail"
		   End If
		Case Else
			Call InventoryReceipts()
			OracleFormWindow("Receipt Header").CloseWindow
			OracleFormWindow("Receipts").CloseWindow
			OracleFormWindow("Navigator").SelectMenu "File->Switch Responsibility..."
			OracleListOfValues("Responsibilities").Select "RS US Payables Clerk"
			OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleButton("Expand Branch").Click
			OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List").Select "          Invoice Batches"
			OracleFormWindow("Navigator").OracleButton("Open").Click
			If OracleFormWindow("Invoice Batches").OracleTextField("Batch Name").Exist(120) Then
				OracleFormWindow("Invoice Batches").OracleTextField("Batch Name").Enter RACK_GetData("Login_Data", "BatchName")
				RACK_ReportEvent "Create Invoice", RACK_GetData("Login_Data", "BatchName") & " is entered as batch name","Pass"
				OracleFormWindow("Invoice Batches").OracleButton("Invoices").Click
				OracleFormWindow("Invoice Workbench").OracleTextField("PO Number").SetFocus
				OracleFormWindow("Invoice Workbench").OracleTextField("PO Number").Enter Environment.Value("POID")
				OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Date").SetFocus
				OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Num").SetFocus
				Environment.Value("InvoiceID") = Invoice()
				OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Num").Enter Environment.Value("InvoiceID")
				RACK_ReportEvent "Create Invoice", Environment.Value("InvoiceID")& " is entered as invoice number" ,"Pass"
				OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Amount").Enter RACK_GetData("Login_Data", "Price")
				RACK_ReportEvent "Create Invoice", RACK_GetData("Login_Data", "Price")& " is entered as amount","Pass"
				OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Amount").SetFocus
				OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Amount").Enter RACK_GetData("Login_Data", "Price")
				OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("PO Line Number").SetFocus
				OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("PO Line Number").OpenDialog
				OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("PO Shipment Number").OpenDialog
				If RACK_GetData("Login_Data", "Tax") = "Tax Accrual" Then
					OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Tax Classification Code").SetFocus
					OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Tax Classification Code").Enter RACK_GetData("Login_Data", "TaxCode")
				End If
				OracleFormWindow("Invoice Workbench").SelectMenu "File->Save"			
				status = OracleStatusLine("OracleStatusLine").GetROProperty("message")
				If  InStr(1,status,"saved") > 0 Then
					RACK_ReportEvent "Create Invoice", "PO match invoice created.Invoice number "& Environment.Value("InvoiceID"),"Pass"
					RACK_ReportEvent "Create Invoice", "PO match invoice has been created successfully ","Screenshot"
				End If
			Else
				RACK_ReportEvent "Create Invoice", "Batch Name  text field does not exist","Fail"
			End If				
   End Select
	If  RACK_GetData("Login_Data", "Validate") = "Yes" Then
		If OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Exist(20) Then
			OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
			OracleFormWindow("Invoice Actions").OracleCheckbox("Validate").Select
			OracleFormWindow("Invoice Actions").OracleButton("OK").Click
			status = OracleFormWindow("Invoice Workbench").OracleTabbedRegion("1 General").OracleTextField("Status|Status").GetROProperty("value")
			'MsgBox(status)
			RACK_ReportEvent " Create Invoice", "Invoice Status "& status ,"Pass"
			RACK_ReportEvent " Create Invoice", "Invoice Status "& status ,"Screenshot"
		Else
			RACK_ReportEvent "Create Invoice", "Actions... 1 button does not exist","Fail"
		End If
	End If
	If  RACK_GetData("Login_Data", "CreateAccounting") = "Yes" Then
		If OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Exist(20) Then
			OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
			OracleFormWindow("Invoice Actions").OracleCheckbox("Create Accounting").Select
			OracleFormWindow("Invoice Actions").OracleRadioGroup("Draft").Select "Final"
			OracleFormWindow("Invoice Actions").OracleButton("OK").Click
			If OracleNotification("Note").Exist(20) Then
				OracleNotification("Note").Approve			
				status = OracleFormWindow("Invoice Workbench").OracleTabbedRegion("1 General").OracleTextField("Status|Accounted").GetROProperty("value")
				'MsgBox(status)
				RACK_ReportEvent " Create Invoice", "Accounted  Status "& status ,"Pass"
				RACK_ReportEvent " Create Invoice", "Accounted  Status "& status ,"Screenshot"
			Else
				RACK_ReportEvent "Create Invoice", "Notification does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Create Invoice", "Actions... 1 button does not exist","Fail"
		End If
	End If
	If  RACK_GetData("Login_Data", "Approve") = "Yes" Then
			If OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Exist(20) Then
				OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
				OracleFormWindow("Invoice Actions").OracleCheckbox("Initiate Approval").Select
				OracleFormWindow("Invoice Actions").OracleButton("OK").Click
				status = OracleFormWindow("Invoice Workbench").OracleTabbedRegion("1 General").OracleTextField("Status|Approval").GetROProperty("value")
				'MsgBox(status)
				RACK_ReportEvent " Create Invoice", "Approval  Status "& status ,"Pass"
				RACK_ReportEvent " Create Invoice", "Approval  Status "& status ,"Screenshot"
            Else
				RACK_ReportEvent "Create Invoice", "Actions... 1 button does not exist","Fail"
			End If
	End If	
End Sub

Sub CancelPO()
If Browser("Login").Page("Oracle Applications Home").Link("RS US Purchasing Staff").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS US Purchasing Staff").Click
	If Browser("Login").Page("Oracle Applications Home_2").Link("Purchase Orders").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home_2").Link("Purchase Orders").Click
		Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
		Browser("Login").Page("Oracle Applications Home_2").Sync
		If OracleFormWindow("Purchase Orders").OracleTextField("Supplier").Exist(120) Then
			OracleFormWindow("Purchase Orders").OracleTextField("Supplier").SetFocus
			OracleFormWindow("Purchase Orders").OracleTextField("Supplier").Enter RACK_GetData("Login_Data", "Supplier")
            OracleFormWindow("Purchase Orders").OracleTextField("Site").SetFocus			
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Type").SetFocus
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Type").Enter "Goods"		
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Item").SetFocus
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Item").Enter RACK_GetData("Login_Data", "item")
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Quantity").SetFocus
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Quantity").Enter RACK_GetData("Login_Data", "Qty")
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Price").Enter RACK_GetData("Login_Data", "Price")
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Need-By").SetFocus
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Need-By").Enter RACK_GetData("Login_Data", "NeedBy")
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Num").SetFocus
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Item_2").SetFocus
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Item_2").Enter RACK_GetData("Login_Data", "item")
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Quantity_2").SetFocus
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Quantity_2").Enter RACK_GetData("Login_Data", "Qty")
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Price_2").Enter RACK_GetData("Login_Data", "Price")
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Need-By_2").SetFocus
			OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Need-By_2").Enter RACK_GetData("Login_Data", "NeedBy")
            OracleFormWindow("Purchase Orders").SelectMenu "File->Save"           	
			PO_ID = OracleFormWindow("Purchase Orders").OracleTextField("PO, Rev").GetROProperty("value")
			'MsgBox(PO_ID)
            RACK_ReportEvent "Create PO", "PO "&PO_ID&" Created succecssfully","Pass"
			If OracleFormWindow("Purchase Orders").OracleButton("Approve...").Exist(20) Then
				OracleFormWindow("Purchase Orders").OracleButton("Approve...").Click
				If OracleFormWindow("Approve Document").OracleTabbedRegion("Approval Details").OracleCheckbox("Approval|Submit for Approval").Exist(20) Then
					OracleFormWindow("Approve Document").OracleTabbedRegion("Approval Details").OracleCheckbox("Approval|Submit for Approval").Select
					OracleFormWindow("Approve Document").OracleButton("OK").Click
					 status = OracleFormWindow("Purchase Orders").OracleTextField("Status").GetROProperty("value")
					RACK_ReportEvent  " PO status", "PO - " & status, "Pass"
					If RACK_GetData("Login_Data", "Transaction") = "cancelline" Then
						OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Num").SetFocus
					End If
					OracleFormWindow("Purchase Orders").SelectMenu "Tools->Cancel"
					OracleFormWindow("Control Document").OracleButton("OK").Click
					OracleNotification("Caution").Approve
					If RACK_GetData("Login_Data", "Transaction") = "cancelline" Then
						If OracleFormWindow("Purchase Orders").OracleTabbedRegion("Lines").OracleTextField("Num").GetROProperty("editable") = "True"  Then
							RACK_ReportEvent  " PO status", "PO line is not cancelled ", "Fail"							
						Else
							RACK_ReportEvent  " PO status", "PO line is cancelled  " , "Pass"
						End If
					Else
						OracleFormWindow("Purchase Orders").CloseWindow
						OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleList("Function List").Select "+  Purchase Orders"
						OracleFormWindow("Navigator").OracleTabbedRegion("Functions").OracleButton("Expand Branch").Click
						OracleFormWindow("Navigator").OracleButton("Open").Click
						OracleFormWindow("Find Purchase Orders").OracleTextField("Number").Enter PO_ID
						OracleFormWindow("Find Purchase Orders").OracleButton("Find").Click
						status = OracleFormWindow("Purchase Order Headers").OracleCheckbox("Cancelled").GetROProperty("selected")
						If status = "True" Then
							RACK_ReportEvent  " PO status", "PO cancelled  " & status, "Pass"
						Else
							RACK_ReportEvent  " PO status", "PO cancel status is   " & status, "Fail"
						End If
					End If
				Else
					RACK_ReportEvent "Cancel PO", "Approve checkbox does not exist","Fail"
				End If
			Else	
				RACK_ReportEvent "Cancel PO", "Approve button does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Cancel PO", "Supplier text field does not exist","Fail"
		End If
	Else
		RACK_ReportEvent "Cancel PO", "Purchase Orders link does not exist","Fail"
	End If
Else
	RACK_ReportEvent "Cancel PO", "RS US Purchasing Staff link does not exist","Fail"
End If
End Sub
Sub HoldInvoice()
 If Browser("Login").Page("Oracle Applications Home").Link("RS US Payables Clerk").Exist(20) Then
			Browser("Login").Page("Oracle Applications Home").Link("RS US Payables Clerk").Click
			If Browser("Login").Page("Oracle Applications Home_2").Link("Invoice Batches").Exist(20) Then
				Browser("Login").Page("Oracle Applications Home_2").Link("Invoice Batches").Click
				Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
				Browser("Login").Page("Oracle Applications Home_2").Sync
				If OracleFormWindow("Invoice Batches").OracleTextField("Batch Name").Exist(120) Then
					OracleFormWindow("Invoice Batches").OracleTextField("Batch Name").Enter RACK_GetData("Login_Data", "BatchName")
					OracleFormWindow("Invoice Batches").OracleButton("Invoices").Click
					If OracleFormWindow("Invoice Workbench").OracleTextField("Trading Partner").Exist(20) Then
						OracleFormWindow("Invoice Workbench").OracleTextField("Trading Partner").Enter RACK_GetData("Login_Data", "Supplier")
						OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Date").SetFocus
						OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Num").SetFocus
						OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Num").Enter Invoice()
						OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Amount").Enter RACK_GetData("Login_Data", "invoicePrice")
						OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Amount").SetFocus
						OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleTextField("Amount").Enter RACK_GetData("Login_Data", "Price")
                        OracleFormWindow("Invoice Workbench").OracleTabbedRegion("2 Lines").OracleButton("Distributions").Click
						If OracleFormWindow("Distributions").OracleTextField("Amount").Exist(20) Then
							OracleFormWindow("Distributions").OracleTextField("Amount").Enter RACK_GetData("Login_Data", "Price")
							OracleFormWindow("Distributions").OracleTextField("Account").SetFocus
							OracleFormWindow("Distributions").OracleTextField("Account").Enter RACK_GetData("Login_Data", "Account")
							OracleFormWindow("Distributions").SelectMenu "File->Save"
							OracleFormWindow("Distributions").CloseWindow       
							If OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Exist(20) Then
								OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
								OracleFormWindow("Invoice Actions").OracleCheckbox("Validate").Select
								OracleFormWindow("Invoice Actions").OracleButton("OK").Click
								status = OracleFormWindow("Invoice Workbench").OracleTabbedRegion("1 General").OracleTextField("Status|Status").GetROProperty("value")
                                If status =  "Needs Revalidation" Then
									reason = OracleFormWindow("Invoice Workbench").OracleTabbedRegion("3 Holds").OracleTextField("Hold Name").GetROProperty("value")
									RACK_ReportEvent " Hold Invoice", "Invoice is in  hold - Reason  " &reason ,"Pass"
								Else
									RACK_ReportEvent " Hold Invoice", "Invoice is not on hold. Invoice status is  " & status ,"Fail"
								End If  	
							Else
								RACK_ReportEvent "Create Invoice", "Actions... 1 button does not exist","Fail"
							End If
						Else
							RACK_ReportEvent "Create Invoice", "Amount  text field does not exist","Fail"
						End If
					Else
						RACK_ReportEvent "Create Invoice", "Trading Partner text field does not exist","Fail"
					End If
				Else
					RACK_ReportEvent "Create Invoice", "Batch Name  text field does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Create Invoice", "Invoice batches link does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Create Invoice", "RS US Payables Clerk link does not exist","Fail"
	   End If
End Sub
Sub ReleaseHold()
	Call HoldInvoice()
	totalamt = OracleFormWindow("Invoice Workbench").OracleTabbedRegion("1 General").OracleTextField("Summary|Total").GetROProperty("value")
	OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Amount").SetFocus
	OracleFormWindow("Invoice Workbench").OracleTextField("Invoice Amount").Enter totalamt
	OracleFormWindow("Invoice Workbench").SelectMenu "File->Save"
	If OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Exist(20) Then
		OracleFormWindow("Invoice Workbench").OracleButton("Actions... 1").Click
		OracleFormWindow("Invoice Actions").OracleCheckbox("Validate").Select
		OracleFormWindow("Invoice Actions").OracleButton("OK").Click
		status = OracleFormWindow("Invoice Workbench").OracleTabbedRegion("1 General").OracleTextField("Status|Status").GetROProperty("value")
		If status =  "Validated" Then
			reason = OracleFormWindow("Invoice Workbench").OracleTabbedRegion("3 Holds").OracleTextField("Release Reason").GetROProperty("value")
			RACK_ReportEvent " Release Hold ", "Invoice is released. Reason  " &reason ,"Pass"
		Else
			RACK_ReportEvent " Release Hold ", "Invoice is on hold. Invoice status is  " & status ,"Fail"
		End If  
	Else
		RACK_ReportEvent "Create Invoice", "Actions... 1 button does not exist","Fail"
	End If  
End Sub
Sub InterOrgShippment()
If Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory User,").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory User,").Click
	If Browser("Login").Page("Oracle Applications Home_2").Link("Inter-organization Transfer").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home_2").Link("Inter-organization Transfer").Click
		Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
		Browser("Login").Page("Oracle Applications Home_2").Sync
		If OracleListOfValues("Organizations").Exist(120) Then
			OracleListOfValues("Organizations").Select RACK_GetData("Login_Data", "Organization")
			If OracleFormWindow("Inter").OracleTextField("Transaction|To Org").Exist(20) Then
				OracleFormWindow("Inter").OracleTextField("Transaction|To Org").SetFocus
				OracleFormWindow("Inter").OracleTextField("Transaction|To Org").OpenDialog
				OracleListOfValues("Transfer Organizations").Select RACK_GetData("Login_Data", "ToOrganization")
				OracleFormWindow("Inter").OracleTextField("Transaction|Type").OpenDialog
				OracleFormWindow("Inter").OracleTextField("Shipment|Number").Enter RACK_GetData("Login_Data", "ShippmentNo")
				OracleFormWindow("Inter").OracleButton("Transaction Lines").Click
				If OracleFormWindow("Intransit Shipment").OracleTextField("Item").Exist(20) Then
					OracleFormWindow("Intransit Shipment").OracleTextField("Item").Enter RACK_GetData("Login_Data", "item")
					OracleFormWindow("Intransit Shipment").OracleTextField("Subinventory").SetFocus 
					OracleFormWindow("Intransit Shipment").OracleTextField("Subinventory").Enter RACK_GetData("Login_Data", "Subinventory")
					OracleFormWindow("Intransit Shipment").OracleTextField("Locator").Enter RACK_GetData("Login_Data", "Locator")
					OracleFormWindow("Intransit Shipment").OracleTextField("Quantity").SetFocus
					OracleFormWindow("Intransit Shipment").OracleTextField("Quantity").Enter RACK_GetData("Login_Data", "Qty")
					OracleFormWindow("Intransit Shipment").OracleButton("Lot / Serial").Click
					If OracleFormWindow("Serial Entry").OracleTextField("Start Serial Number").Exist(20) Then
						OracleFormWindow("Serial Entry").OracleTextField("Start Serial Number").Enter RACK_GetData("Login_Data", "serialnumber")
						OracleFormWindow("Serial Entry").OracleButton("Done").Click
						OracleFormWindow("Intransit Shipment").SelectMenu "File->Save"
						status = OracleStatusLine("OracleStatusLine").GetROProperty("message")
						If  InStr(1,status,"saved") > 0 Then
							RACK_ReportEvent " Inter Org Shippment", "Inter Org Shippment completed. Shippment Number  " &  RACK_GetData("Login_Data", "ShippmentNo") ,"Pass"
						Else
							RACK_ReportEvent " Inter Org Shippment", "Inter Org Shippment not completed." ,"Fail"
						End If
					Else					
						RACK_ReportEvent "Inter Org Shippment ", "Start Serial Number text field does not exist","Fail"
					End If
				Else
					RACK_ReportEvent "Inter Org Shippment ", "Item text field does not exist","Fail"
				End If
			Else
				RACK_ReportEvent "Inter Org Shippment ", "Transaction|To Org text field does not exist","Fail"
			End If
		Else
			RACK_ReportEvent "Inter Org Shippment ", "Organizations list does not exist","Fail"
		End If
	Else
		RACK_ReportEvent "Inter Org Shippment ", "Inter-organization Transfer link does not exist","Fail"
	End If
Else
	RACK_ReportEvent "Inter Org Shippment ", "RS US  Inventory User link does not exist","Fail"
End If
End Sub

Sub GLInquiry()
If Browser("Login").Page("Oracle Applications Home").Link("RS US GL User").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS US GL User").Click
	If Browser("Login").Page("Oracle Applications Home_2").Link("Account").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home_2").Link("Account").Click
		Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
		Browser("Login").Page("Oracle Applications Home_2").Sync
		If OracleFormWindow("Account Inquiry").OracleTextField("Accounting Periods|From").Exist(120) Then
			OracleFormWindow("Account Inquiry").OracleTextField("Accounting Periods|From").Enter RACK_GetData("Login_Data", "Period")
			OracleFormWindow("Account Inquiry").OracleTabbedRegion("Primary Balance Type").OracleRadioGroup("Actual").Select "Actual"
			OracleFormWindow("Account Inquiry").OracleTextField("Account").SetFocus
			OracleFlexWindow("Find Accounts").OracleTextField("Company").Enter RACK_GetData("Login_Data", "Company")
			OracleFlexWindow("Find Accounts").OracleTextField("Location").Enter RACK_GetData("Login_Data", "Location")
			OracleFlexWindow("Find Accounts").OracleTextField("Account").Enter RACK_GetData("Login_Data", "Account")
			OracleFlexWindow("Find Accounts").OracleTextField("Team").Enter RACK_GetData("Login_Data", "Account")
			OracleFlexWindow("Find Accounts").OracleTextField("Business Unit").Enter RACK_GetData("Login_Data", "BusinessUnit")
			OracleFlexWindow("Find Accounts").OracleTextField("Department").Enter RACK_GetData("Login_Data", "Department")
			OracleFlexWindow("Find Accounts").OracleTextField("Product").Enter RACK_GetData("Login_Data", "Product")
			OracleFlexWindow("Find Accounts").OracleTextField("Future").Enter RACK_GetData("Login_Data", "Future")
			OracleFlexWindow("Find Accounts").Approve
			If OracleFormWindow("Account Inquiry").OracleButton("Show Balances").Exist(20) Then
				OracleFormWindow("Account Inquiry").OracleButton("Show Balances").Click
				balance = OracleFormWindow("Detail Balances").OracleTextField("YTD").GetROProperty("value")
				RACK_ReportEvent " GL Inquiry", "Account balance "& balance ,"Pass"
			Else
					RACK_ReportEvent " GL Inquiry ", "Show Balances button does not exist","Fail"
			End If
		Else
			RACK_ReportEvent " GL Inquiry ", "Accounting Periods text field does not exist","Fail"
		End If
	Else
		RACK_ReportEvent " GL Inquiry ", "Account link does not exist","Fail"
	End If
Else
	RACK_ReportEvent " GL Inquiry ", "RS US GL User link does not exist","Fail"
End If
End Sub

Sub CancelPO_Requisition()
OracleFormWindow("AutoCreate to Purchase").OracleTabbedRegion("Lines").OracleTextField("Item").SetFocus
OracleFormWindow("AutoCreate to Purchase").SelectMenu "Tools->Cancel"
OracleFormWindow("Control Document").OracleCheckbox("Cancel Requisition").Select
OracleFormWindow("Control Document").OracleButton("OK").Click
OracleNotification("Caution").Approve
status = OracleFormWindow("AutoCreate to Purchase").OracleTextField("Status").GetROProperty("value")
If  InStr(1,status,"Closed") > 0 Then
	RACK_ReportEvent  " Cancel PO Requisition", "PO - status  " & status, "Pass"
Else
		RACK_ReportEvent  " Cancel PO Requisition", "PO - status  " & status, "Fail"
End If
End Sub
Sub OnHandQuantity()
If Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory User,").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory User,").Click
	If Browser("Login").Page("Oracle Applications Home_2").Link("On-hand Quantity").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home_2").Link("On-hand Quantity").Click
		Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
		Browser("Login").Page("Oracle Applications Home_2").Sync
		If OracleListOfValues("Organizations").Exist(120) Then
			OracleListOfValues("Organizations").Select RACK_GetData("Login_Data", "Organization")
			OracleFormWindow("Query Material").OracleTabbedRegion("Item").OracleTextField("Item|Item / Revision").Enter RACK_GetData("Login_Data", "item")
			OracleFormWindow("Query Material").OracleButton("Find").Click
			If OracleFormWindow("Material Workbench").OracleTabbedRegion("Quantity").OracleButton("Availability").Exist(20) Then
				OracleFormWindow("Material Workbench").OracleTabbedRegion("Quantity").OracleButton("Availability").Click
				If OracleFormWindow("Availability").OracleTextField("Available to Reserve|Primary").Exist Then
					qty = OracleFormWindow("Availability").OracleTextField("Available to Reserve|Primary").GetROProperty("value")
					RACK_ReportEvent  " On Hand Quantity", "Available to Reserve quantity for item  "&RACK_GetData("Login_Data", "item")&" is " & qty, "Pass"
					qty = OracleFormWindow("Availability").OracleTextField("Total Quantity|Primary").GetROProperty("value")
					RACK_ReportEvent  " On Hand Quantity", "Total Quantity for item "&RACK_GetData("Login_Data", "item")&" is "  & qty, "Pass"
				Else
					RACK_ReportEvent  "  On Hand Quantity ", "  Available to Reserve text field does not exist","Fail"
				End If
			Else
				RACK_ReportEvent  "  On Hand Quantity ", "  Item text field does not exist","Fail"
			End If
		Else
			RACK_ReportEvent  "  On Hand Quantity ", "  Organizations list does not exist","Fail"
		End If
	Else
		RACK_ReportEvent  "  On Hand Quantity ", "On-hand Quantity  link does not exist","Fail"
	End If
Else
	RACK_ReportEvent  "  On Hand Quantity ", "RS US Inventory User  link does not exist","Fail"
End If
End Sub
Sub MaterialInquiry()
If Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory User,").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS US Inventory User,").Click
	If Browser("Login").Page("Oracle Applications Home_2").Link("Material Transactions").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home_2").Link("Material Transactions").Click
		Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
		Browser("Login").Page("Oracle Applications Home_2").Sync
		If OracleListOfValues("Organizations").Exist(120) Then			
			OracleListOfValues("Organizations").Select RACK_GetData("Login_Data", "Organization")
			If OracleFormWindow("Find Material Transactions").OracleTextField("Transaction Dates").Exist(20) Then
				OracleFormWindow("Find Material Transactions").OracleTextField("Transaction Dates").Enter ""
				OracleFormWindow("Find Material Transactions").OracleTextField("To Date: Transaction Dates").Enter ""
				OracleFormWindow("Find Material Transactions").OracleTextField("Item").Enter RACK_GetData("Login_Data", "item")
				OracleFormWindow("Find Material Transactions").OracleButton("Find").Click
				If OracleFormWindow("Material Transactions").OracleTabbedRegion("Transaction Type").Exist(50) Then
						RACK_ReportEvent  "Material Inquiry" , "Material Inquiry done sucessfully for the item  "&RACK_GetData("Login_Data", "item") , "Pass"
				Else
						RACK_ReportEvent  "Material Inquiry" , "Transaction Type Tabbed Region does not exist", "Fail"	
				End if
			Else
				RACK_ReportEvent  "Material Inquiry ", "Transaction Dates text field does not exist","Fail"
			End If
		Else
			RACK_ReportEvent  "Material Inquiry ", "Organizations list does not exist","Fail"
		End If
	Else
		RACK_ReportEvent  "Material Inquiry ", "Material Transactions link does not exist","Fail"
	End If
Else
	RACK_ReportEvent  "Material Inquiry ", "RS US Inventory User  link does not exist","Fail"
End If
End Sub
Sub CreateEmployee()
If Browser("Login").Page("Oracle Applications Home").Link("RS US HRMS Manager, Rackspace").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS US HRMS Manager, Rackspace").Click
	If Browser("Login").Page("Oracle Applications Home_2").Link("Enter and Maintain").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home_2").Link("Enter and Maintain").Click
		Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
		Browser("Login").Page("Oracle Applications Home_2").Sync
		If OracleFormWindow("Find Person").OracleButton("New").Exist(120) Then
			OracleFormWindow("Find Person").OracleButton("New").Click
			If OracleFormWindow("People").OracleTextField("Name|Last").Exist(20) Then
				OracleFormWindow("People").OracleTextField("Name|Last").Enter RACK_GetData("Login_Data", "Name")
				OracleFormWindow("People").OracleList("Gender").Select RACK_GetData("Login_Data", "Gender")
				OracleFormWindow("People").OracleList("Action").Select "Create Employment"				
				ssn = CStr(Int((999-100+1)*Rnd+100)) &"-"&CStr(Int((99-10+1)*Rnd+10))&"-"&CStr(Int((9999-1000+1)*Rnd+1000))
				OracleFormWindow("People").OracleTextField("Identification|Social").Enter ssn
				OracleFormWindow("People").OracleTabbedRegion("Personal").OracleTextField("Birth Date").Enter RACK_GetData("Login_Data", "DOB")
				OracleFormWindow("People").SelectMenu "File->Save"
				'OracleFormWindow("People").OracleButton("Assignment").Click
				status = OracleStatusLine("OracleStatusLine").GetROProperty("message")

				If  InStr(1,status,"saved") > 0 Then
					RACK_ReportEvent " Create Employee", "Employee created successfully  " ,"Pass"
				Else
					RACK_ReportEvent " Create Employee", "Employee not created successfully" ,"Fail"
				End If
			Else
					RACK_ReportEvent  " Create Employee", "Name text field does not exist","Fail"
			End If
		Else
				RACK_ReportEvent  " Create Employee", "New  Button does not exist","Fail"
		End If
	Else
			RACK_ReportEvent  " Create Employee", "Enter and Maintain  link does not exist","Fail"
	End If
Else
	RACK_ReportEvent  " Create Employee", "RS US HRMS Manager  link does not exist","Fail"
End If
End Sub
Sub ModifyEmployee()
CreateEmployee()
If OracleFormWindow("People").OracleTextField("Name|Last").Exist(20) Then	
	OracleFormWindow("People").OracleTabbedRegion("Personal").OracleTextField("Birth Date").InvokeSoftkey "ENTER QUERY"
	OracleFormWindow("People").OracleTextField("Name|Last").Enter RACK_GetData("Login_Data", "Name")
	OracleFormWindow("People").OracleTextField("Name|Last").InvokeSoftkey "EXECUTE QUERY"
	If OracleFormWindow("People").OracleTabbedRegion("Personal").OracleTextField("Birth Date").Exist(20) Then
		OracleFormWindow("People").OracleTabbedRegion("Personal").OracleTextField("Birth Date").Enter RACK_GetData("Login_Data", "NewDOB")
		
		If OracleFormWindow("Choose an option").OracleButton("Correction").Exist(20) Then
			OracleFormWindow("Choose an option").OracleButton("Correction").Click
			OracleFormWindow("People").SelectMenu "File->Save"
			If OracleFormWindow("People").Exist(20) Then
				status = OracleStatusLine("OracleStatusLine").GetROProperty("message")
				If  InStr(1,status,"saved") > 0 Then
					RACK_ReportEvent " Modify Employee", "Employee's DOB has been modified " ,"Pass"
				Else
					RACK_ReportEvent " Modify Employee", "Employee's DOB has not been modified" ,"Fail"
				End If
			Else	
				RACK_ReportEvent " Modify Employee", "People form does not exist" ,"Fail"
			End If
		Else
			RACK_ReportEvent " Modify Employee", "Correction button does not exist" ,"Fail"
		End If
	Else
		RACK_ReportEvent " Modify Employee", "Birth Date text field does not exist" ,"Fail"
	End If
Else
    RACK_ReportEvent " Modify Employee", "Name|Last text field does not exist" ,"Fail"
End If
End Sub
Sub Notification()
If Browser("Login").Page("Oracle Applications Home").Link("RS US Purchasing Approver").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS US Purchasing Approver").Click
	If Browser("Login").Page("Oracle Applications Home_2").Link("Notifications").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home_2").Link("Notifications").Click
'		Browser("Login").Page("Oracle Workflow: Notifications").RefreshObject
		Browser("Login").Page("Oracle Workflow: Notifications").WebList("NtfView").Select "Open Notifications"
		Browser("Login").Page("Oracle Workflow: Notifications").WebButton("Go").Click
		Browser("Login").Page("Oracle Workflow: Notifications").Link("html tag:=A","text:=Rackspace US - Operating Unit - Standard Purchase Order "&Environment.Value("POID") &" for .*").Click	
		'Browser("Login").Page("Oracle Workflow: Notifications").WebCheckBox("N25:selected:0").Set "ON"
		'Browser("Login").Page("Oracle Workflow: Notifications").WebButton("Open").Click
		
		Select Case RACK_GetData("Login_Data", "Action")
			Case "Reassign"
				Browser("Login").Page("Notification Details").WebButton("Reassign").Click
				Browser("Login").Page("Reassign Notifications").WebList("wfUserType1").Select "Employee"
				Browser("Login").Page("Reassign Notifications").WebEdit("wfUserName1").Set RACK_GetData("Login_Data", "Name")
				Browser("Login").Page("Reassign Notifications").WebButton("Submit").Click
				If NOT Browser("Login").Page("Oracle Workflow: Notifications").Link("html tag:=A","text:=Rackspace US - Operating Unit - Standard Purchase Order "&Environment.Value("POID") &" for .*").Exist(2) Then
					RACK_ReportEvent  "  Notification ", "PO "&Environment.Value("POID")&" has been reassigned","Pass"
				Else
					RACK_ReportEvent  "  Notification ", "PO "&Environment.Value("POID")&" has not been reassigned","Pass"
				End If				
            Case "Forward"
				Browser("Login").Page("Notification Details").WebList("wfUserType1").Select "Employee"
				Browser("Login").Page("Notification Details").WebEdit("wfUserName1").Set RACK_GetData("Login_Data", "Name")
				Browser("Login").Page("Notification Details").WebButton("Forward").Click
				If NOT Browser("Login").Page("Oracle Workflow: Notifications").Link("html tag:=A","text:=Rackspace US - Operating Unit - Standard Purchase Order "&Environment.Value("POID") &" for .*").Exist(2) Then
					RACK_ReportEvent  "  Notification ", "PO "&Environment.Value("POID")&" has been forwarded","Pass"
				Else
					RACK_ReportEvent  "  Notification ", "PO "&Environment.Value("POID")&" has not been forwarded","Pass"
				End If				
			Case "Approve And Forward"
				Browser("Login").Page("Notification Details").WebList("wfUserType1").Select "Employee"
				Browser("Login").Page("Notification Details").WebEdit("wfUserName1").Set RACK_GetData("Login_Data", "Name")
				Browser("Login").Page("Notification Details").WebButton("Approve And Forward").Click
				If NOT Browser("Login").Page("Oracle Workflow: Notifications").Link("html tag:=A","text:=Rackspace US - Operating Unit - Standard Purchase Order "&Environment.Value("POID") &" for .*").Exist(2) Then
					RACK_ReportEvent  "  Notification ", "PO "&Environment.Value("POID")&" has been approved and forwarded","Pass"
				Else
					RACK_ReportEvent  "  Notification ", "PO "&Environment.Value("POID")&" has not been approved and forwarded","Pass"
				End If				
			Case "Reject"
				Browser("Login").Page("Notification Details").WebButton("Reject").Click
				If NOT Browser("Login").Page("Oracle Workflow: Notifications").Link("html tag:=A","text:=Rackspace US - Operating Unit - Standard Purchase Order "&Environment.Value("POID") &" for .*").Exist(2) Then
					RACK_ReportEvent  "  Notification ", "PO "&Environment.Value("POID")&" has been rejected","Pass"
				Else
					RACK_ReportEvent  "  Notification ", "PO "&Environment.Value("POID")&" has not been rejected","Pass"
				End If				
			Case "Approve"
				Browser("Login").Page("Notification Details").WebButton("Approve").Click
				If NOT Browser("Login").Page("Oracle Workflow: Notifications").Link("html tag:=A","text:=Rackspace US - Operating Unit - Standard Purchase Order "&Environment.Value("POID") &" for .*").Exist(2) Then
					RACK_ReportEvent  "  Notification ", "PO "&Environment.Value("POID")&" has been approved ","Pass"
                Else
					RACK_ReportEvent  "  Notification ", "PO "&Environment.Value("POID")&" has not been approved ","Pass"
				End If				
		End Select

	Else
		RACK_ReportEvent  "  Notification ", "Notifications link does not exist","Fail"
	End If
Else
	RACK_ReportEvent  "  Notification ", "RS US Purchasing Approver link does not exist","Fail"
End If
End Sub
Sub Create_AssignAssets()
If Browser("Login").Page("Oracle Applications Home").Link("RS US Project Costing").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS US Project Costing").Click
	If Browser("Login").Page("Oracle Applications Home_2").Link("Projects").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home_2").Link("Projects").Click
		Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
		Browser("Login").Page("Oracle Applications Home_2").Sync
		If OracleFormWindow("Find Projects").OracleList("Project|Search For").Exist(120) Then
			OracleFormWindow("Find Projects").OracleList("Project|Search For").Select "Projects"
			OracleFormWindow("Find Projects").OracleTextField("Project|Number").Enter RACK_GetData("Login_Data", "ProjectNumber")
			OracleFormWindow("Find Projects").OracleButton("Find").Click
			If OracleFormWindow("Projects, Templates Summary").OracleButton("Open").Exist(20) Then
				OracleFormWindow("Projects, Templates Summary").OracleButton("Open").Click
				If JavaWindow("Oracle Applications -").JavaInternalFrame("Projects, Templates").JavaSlider("LWScrollbar").Exist(20) Then
					JavaWindow("Oracle Applications -").JavaInternalFrame("Projects, Templates").JavaSlider("LWScrollbar").Drag(2)
					OracleFormWindow("Projects, Templates").OracleTextField("Option Name").SetFocus
					OracleFormWindow("Projects, Templates").OracleButton("Detail").Click
					OracleFormWindow("Projects, Templates").OracleTextField("Option Name_2").SetFocus
					OracleFormWindow("Projects, Templates").OracleButton("Detail").Click
					OracleFormWindow("Asset").SelectMenu "File->New"
					If OracleFormWindow("Asset").OracleTextField("Asset  Name").Exist(20) Then
						OracleFormWindow("Asset").OracleTextField("Asset  Name").Enter RACK_GetData("Login_Data", "Name")
						OracleFormWindow("Asset").OracleTextField("Description").Enter RACK_GetData("Login_Data", "Description")
						OracleFormWindow("Asset").OracleTextField("Asset Category").Enter RACK_GetData("Login_Data", "Category")
						OracleFlexWindow("Asset Category").OracleTextField("MAJOR").OpenDialog
						OracleListOfValues("MAJOR").Select RACK_GetData("Login_Data", "Major")
						OracleFlexWindow("Asset Category").OracleTextField("MINOR").OpenDialog
						OracleListOfValues("MINOR").Select RACK_GetData("Login_Data", "Minor")
						OracleFlexWindow("Asset Category").Approve
						OracleFormWindow("Asset").OracleTextField("Asset Key").OpenDialog
						OracleFormWindow("Asset").OracleTextField("Location").OpenDialog
						OracleFlexWindow("Location").OracleTextField("LOCATION").OpenDialog
						OracleListOfValues("LOCATION").Select RACK_GetData("Login_Data", "Location")
						OracleFlexWindow("Location").OracleTextField("SITES").OpenDialog
						OracleFlexWindow("Location").Approve
						OracleFormWindow("Asset").SelectMenu "File->Save"
						OracleFormWindow("Asset").CloseWindow
						OracleFormWindow("Projects, Templates").OracleTextField("Option Name").SetFocus
						OracleFormWindow("Projects, Templates").OracleButton("Detail").Click
						If OracleFormWindow("Asset Assignments").OracleTextField("Asset Name").Exist(20) Then
							OracleFormWindow("Asset Assignments").OracleTextField("Asset Name").Enter RACK_GetData("Login_Data", "Name")
							OracleFormWindow("Asset Assignments").SelectMenu "File->Save"
						Else
							RACK_ReportEvent  "  Create Assign Assets ", " Asset  Name text field does not exist","Fail"
						End If
					Else
						RACK_ReportEvent  "  Create Assign Assets ", " Asset  Name text field does not exist","Fail"
					End If
				Else
					RACK_ReportEvent  "  Create Assign Assets ", " LWScrollbar does not exist","Fail"
				End If
			Else
				RACK_ReportEvent  "  Create Assign Assets ", " Open button does not exist","Fail"
			End If
		Else
			RACK_ReportEvent  "  Create Assign Assets ", "Project|Search For list does not exist","Fail"
		End If
	Else
		RACK_ReportEvent  "  Create Assign Assets ", "Projects link does not exist","Fail"
	End If
Else
	RACK_ReportEvent  "  Create Assign Assets ", "RS US Project Costing link does not exist","Fail"
End If
End Sub
Sub ViewRequisition()
If Browser("Login").Page("Oracle Applications Home").Link("RS US iProcurement Requestor").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS US iProcurement Requestor").Click
	If Browser("Login").Page("Oracle iProcurement: Shop").Link("Requisitions").Exist(20) Then
		Browser("Login").Page("Oracle iProcurement: Shop").Link("Requisitions").Click
		If Browser("Login").Page("Oracle iProcurement: Requisitions").WebList("ViewsPopList").Exist(20) Then
			Browser("Login").Page("Oracle iProcurement: Requisitions").WebList("ViewsPopList").Select "My Group's Requisitions"
			Browser("Login").Page("Oracle iProcurement: Requisitions").WebButton("Go").Click
			If Browser("Login").Page("Oracle iProcurement: Requisitions").WebRadioGroup("N19:selected").Exist(20) Then
				Browser("Login").Page("Oracle iProcurement: Requisitions").WebRadioGroup("N19:selected").Select "0"
'				status = Browser("Login").Page("Oracle iProcurement: Requisitions").WebRadioGroup("N19:selected").GetROProperty("Checked")
				Browser("Login").Page("Oracle iProcurement: Requisitions").Link("44581").Click
				status = Browser("Login").Page("Oracle iProcurement: Requisitions").WebElement("Requisition 44581").GetROProperty("innertext")
				If  InStr(1,status,"Requisition") > 0	Then
					RACK_ReportEvent " View Requisition", "Requisition has been viewed " ,"Pass"
				Else
					RACK_ReportEvent " View Requisition", "Requisition has not been viewed " ,"Fail"
				End If	
			Else
				RACK_ReportEvent  "View Requisition ", " link does not exist","Fail"
			End If
		Else
			RACK_ReportEvent  "View Requisition ", " link does not exist","Fail"
		End If
	Else
		RACK_ReportEvent  "View Requisition ", " link does not exist","Fail"
	End If
Else
	RACK_ReportEvent  "View Requisition ", " link does not exist","Fail"
End If
End Sub
Sub UserCreation()
If Browser("Login").Page("Oracle Applications Home").Link("RS User Administrator").Exist(20) Then
	Browser("Login").Page("Oracle Applications Home").Link("RS User Administrator").Click
	If Browser("Login").Page("Oracle Applications Home_2").Link("Define").Exist(20) Then
		Browser("Login").Page("Oracle Applications Home_2").Link("Define").Click
		Browser("Oracle Applications R12").Page("Oracle Applications R12").Sync
		Browser("Login").Page("Oracle Applications Home_2").Sync
		If OracleFormWindow("Users").OracleTextField("User Name").Exist(120) Then
			OracleFormWindow("Users").OracleTextField("User Name").Enter RACK_GetData("Login_Data", "Name" )
			OracleFormWindow("Users").OracleTextField("Password").Enter RACK_GetData("Login_Data", "password" )
			OracleFormWindow("Users").OracleTextField("Password").Enter RACK_GetData("Login_Data", "password" )
			OracleFormWindow("Users").OracleTabbedRegion("Direct Responsibilities").OracleTextField("Responsibility").SetFocus
			OracleFormWindow("Users").OracleTabbedRegion("Direct Responsibilities").OracleTextField("Responsibility").Enter RACK_GetData("Login_Data", "Responsibility" )
			OracleFormWindow("Users").SelectMenu "File->Save"
			status = OracleStatusLine("OracleStatusLine").GetROProperty("message")
				If  InStr(1,status,"saved") > 0 Then
					RACK_ReportEvent " User Creation", "User has been created and responsibility has been added " ,"Pass"
				Else
					RACK_ReportEvent " User Creation", "User has not been created" ,"Fail"
				End If
		End If
	End If
End If
End Sub

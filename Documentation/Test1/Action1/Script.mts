Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Shop").Link("Shopping Lists").Click @@ hightlight id_;_Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Shop").Link("Shopping Lists")_;_script infofile_;_ZIP::ssf1.xml_;_
Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Shop_2").WebList("ShoppingListName").Select "Test HK Template2" @@ hightlight id_;_Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Shop 2").WebList("ShoppingListName")_;_script infofile_;_ZIP::ssf2.xml_;_
Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Shop_2").WebButton("Go").Click @@ hightlight id_;_Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Shop 2").WebButton("Go")_;_script infofile_;_ZIP::ssf3.xml_;_
Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Shop_3").WebButton("Add to Cart").Click @@ hightlight id_;_Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Shop 3").WebButton("Add to Cart")_;_script infofile_;_ZIP::ssf4.xml_;_
Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Shop_3").WebButton("View Cart and Checkout").Click @@ hightlight id_;_Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Shop 3").WebButton("View Cart and Checkout")_;_script infofile_;_ZIP::ssf5.xml_;_
Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Checkout").WebButton("Checkout").Click @@ hightlight id_;_Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Checkout").WebButton("Checkout")_;_script infofile_;_ZIP::ssf6.xml_;_
Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Checkout_2").WebButton("Next").Click @@ hightlight id_;_Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Checkout 2").WebButton("Next")_;_script infofile_;_ZIP::ssf7.xml_;_
Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Checkout_3").WebButton("Next").Click @@ hightlight id_;_Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Checkout 3").WebButton("Next")_;_script infofile_;_ZIP::ssf8.xml_;_
Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Checkout_4").WebButton("Submit").Click @@ hightlight id_;_Browser("Oracle iProcurement: Shop").Page("Oracle iProcurement: Checkout 4").WebButton("Submit")_;_script infofile_;_ZIP::ssf9.xml_;_
Browser("Oracle iProcurement: Shop").Page("Confirmation").WebButton("Continue Shopping").Click @@ hightlight id_;_Browser("Oracle iProcurement: Shop").Page("Confirmation").WebButton("Continue Shopping")_;_script infofile_;_ZIP::ssf10.xml_;_



OracleFormWindow("Requisition Template").OracleTextField("Template").Enter "test" @@ hightlight id_;_224_;_script infofile_;_ZIP::ssf11.xml_;_
OracleFormWindow("Requisition Template").OracleButton("Copy...").Click @@ hightlight id_;_248_;_script infofile_;_ZIP::ssf12.xml_;_
OracleFormWindow("Base Document").OracleTextField("Type").Enter "Purchase%order" @@ hightlight id_;_291_;_script infofile_;_ZIP::ssf13.xml_;_
OracleFormWindow("Base Document").OracleTextField("Number").Enter "%%" @@ hightlight id_;_293_;_script infofile_;_ZIP::ssf14.xml_;_
OracleListOfValues("Purchase Orders").Select "10851"
OracleFormWindow("Base Document").OracleButton("OK").Click @@ hightlight id_;_294_;_script infofile_;_ZIP::ssf15.xml_;_
OracleFormWindow("Requisition Template").SelectMenu "File->Save" @@ hightlight id_;_136_;_script infofile_;_ZIP::ssf16.xml_;_
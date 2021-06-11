SystemUtil.CloseProcessByName "iexplore.exe"
SystemUtil.Run "iexplore.exe", "https://advantageonlineshopping.com"
For sendFeedback = 1 To 4 Step 1

	Browser("Advantage Shopping").Page("Advantage Shopping").Link("Link").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").Link("Link")_;_script infofile_;_ZIP::ssf1.xml_;_
	Browser("Advantage Shopping").Page("Advantage Shopping").WebList("categoryListboxContactUs").Select DataTable("categoryListboxContactUs_Item", dtGlobalSheet) @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebList("categoryListboxContactUs")_;_script infofile_;_ZIP::ssf2.xml_;_
	Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("ProductHP Chromebook 14").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("ProductHP Chromebook 14")_;_script infofile_;_ZIP::ssf3.xml_;_
	Browser("Advantage Shopping").Page("Advantage Shopping").WebList("productListboxContactUs").Select DataTable("productListboxContactUs_Item", dtGlobalSheet) @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebList("productListboxContactUs")_;_script infofile_;_ZIP::ssf4.xml_;_
	Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("ProductHP Chromebook 14").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("ProductHP Chromebook 14")_;_script infofile_;_ZIP::ssf5.xml_;_
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("emailContactUs").Set DataTable("emailContactUs_Text", dtGlobalSheet) @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("emailContactUs")_;_script infofile_;_ZIP::ssf6.xml_;_
	Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("supportCover").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("supportCover")_;_script infofile_;_ZIP::ssf7.xml_;_
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("subjectTextareaContactUs").Set DataTable("subjectTextareaContactUs_Text", dtGlobalSheet) @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("subjectTextareaContactUs")_;_script infofile_;_ZIP::ssf8.xml_;_
	Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("* Subject subjectTextareahiSEN").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("* Subject subjectTextareahiSEN")_;_script infofile_;_ZIP::ssf9.xml_;_
	Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("send_btnundefined").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("send btnundefined")_;_script infofile_;_ZIP::ssf10.xml_;_
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("emailContactUs").Set DataTable("emailContactUs_Text1", dtGlobalSheet) @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("emailContactUs")_;_script infofile_;_ZIP::ssf11.xml_;_
	Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("* Subject subjectTextareahiSEN").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("* Subject subjectTextareahiSEN")_;_script infofile_;_ZIP::ssf12.xml_;_
	Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("send_btnundefined").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("send btnundefined")_;_script infofile_;_ZIP::ssf13.xml_;_
	Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Thank you for contacting").Check CheckPoint("Thank you for contacting Advantage support.") @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Thank you for contacting")_;_script infofile_;_ZIP::ssf14.xml_;_
	Browser("Advantage Shopping").Page("Advantage Shopping").Link("CONTINUE SHOPPING").Click @@ hightlight id_;_Browser("Advantage Shopping").Page("Advantage Shopping").Link("CONTINUE SHOPPING")_;_script infofile_;_ZIP::ssf15.xml_;_
	
Next
SystemUtil.CloseProcessByName "iexplore.exe"

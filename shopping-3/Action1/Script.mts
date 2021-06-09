wait 10
Browser("Advantage Shopping").Page("Advantage Shopping").Link("CONTACT US").Click @@ script infofile_;_ZIP::ssf1.xml_;_
wait 10
Browser("Advantage Shopping").Page("Advantage Shopping").WebList("categoryListboxContactUs").Select DataTable("categoryListboxContactUs_Item", dtGlobalSheet) @@ script infofile_;_ZIP::ssf2.xml_;_
Browser("Advantage Shopping").Page("Advantage Shopping").WebList("productListboxContactUs").Select DataTable("productListboxContactUs_Item", dtGlobalSheet) @@ script infofile_;_ZIP::ssf3.xml_;_
Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("emailContactUs").Set DataTable("emailContactUs_Text", dtGlobalSheet) @@ script infofile_;_ZIP::ssf4.xml_;_
Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("subjectTextareaContactUs").Set DataTable("subjectTextareaContactUs_Text", dtGlobalSheet) @@ script infofile_;_ZIP::ssf5.xml_;_
Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("send_btnundefined").Click @@ script infofile_;_ZIP::ssf6.xml_;_
Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Thank you for contacting").Check CheckPoint("Thank you for contacting Advantage support.") @@ script infofile_;_ZIP::ssf7.xml_;_
Browser("Advantage Shopping").Page("Advantage Shopping").Link("CONTINUE SHOPPING").Click @@ script infofile_;_ZIP::ssf8.xml_;_

SystemUtil.Run "msedge.exe","https://www.advantageonlineshopping.com/"
Browser("Advantage Shopping").Page("Advantage Shopping").Sync @@ script infofile_;_ZIP::ssf5.xml_;_
Browser("Advantage Shopping").Page("Advantage Shopping").Link("SpeakersCategory").Click @@ script infofile_;_ZIP::ssf7.xml_;_
Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("$44.99").Click @@ script infofile_;_ZIP::ssf9.xml_;_
Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("save_to_cart").Click @@ script infofile_;_ZIP::ssf11.xml_;_
'verify the checkout button is visible to click
If Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("check_out_btn").Exist(10) Then @@ script infofile_;_ZIP::ssf14.xml_;_
	Reporter.ReportEvent micPass,"checkout button" , "test passed"
Else
	Reporter.ReportEvent micPass,"checkout button" , "test failed"
End If


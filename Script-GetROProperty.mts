varA=WpfWindow("OpenText MyFlight Sample").WpfButton("OK").getROProperty("enabled")
print varA

If varA=False Then
	'msgBox "ok button is disabled"
	Reporter.ReportEvent micPass, "ok button status","ok button disabled on app launch"
else
'msgBox "ok button is enabled"
Reporter.ReportEvent micFail, "ok button status","ok button enabled app launch"
End If



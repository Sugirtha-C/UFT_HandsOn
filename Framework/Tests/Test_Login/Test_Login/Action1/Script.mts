' Main Test Script for Login
Sub TestLogin()
    ' Read data from DataTable
    Dim url, username, password
    url = DataTable("URL", dtGlobalSheet)
    username = DataTable("Username", dtGlobalSheet)
    password = DataTable("Password", dtGlobalSheet)
    ' Open Browser and navigate to URL
    Call OpenBrowser(url)
     Call CaptureScreenshot("OpenBrowser")
    ' Perform Login
    Call PerformLogin(username, password)
    Call CaptureScreenshot("AfterLoginAttempt")
    ' Verify Login
    If VerifyLogin() Then
        Reporter.ReportEvent micPass, "TestLogin", "TestLogin passed"
    Else
        Reporter.ReportEvent micFail, "TestLogin", "TestLogin failed"
    End If
    Call CaptureScreenshot("LoginVerification")
    ' Close Browser
    Call CloseBrowser()
End Sub
' Execute the test
Call TestLogin()


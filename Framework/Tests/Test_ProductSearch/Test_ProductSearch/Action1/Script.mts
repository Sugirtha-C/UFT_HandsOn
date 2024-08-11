' Main Test Script for Product Search
Sub TestProductSearch()
    ' Read data from DataTable
    Dim url, username, password, productName
    url = DataTable("URL", dtGlobalSheet)
    username = DataTable("Username", dtGlobalSheet)
    password = DataTable("Password", dtGlobalSheet)
    productName = DataTable("ProductName", dtGlobalSheet)
    ' Open Browser and navigate to URL
    Call OpenBrowser(url)
    ' Perform Login
    Call PerformLogin(username, password)
    ' Verify Login
    If  VerifyLogin() Then
        ' Search for Product
        Call SearchProduct(productName)
    End If
    ' Close Browser
    Call CloseBrowser()
End Sub
' Execute the test
Call TestProductSearch()


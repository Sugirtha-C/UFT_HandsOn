	' Main Test Script for Checkout
Sub TestCheckout()
    ' Read data from DataTable
    Dim url, username, password, productName
    url = DataTable("URL", dtGlobalSheet)
    username = DataTable("Username", dtGlobalSheet)
    password = DataTable("Password", dtGlobalSheet)
' Main Test Script for Product Search
    productName = DataTable("ProductName", dtGlobalSheet) 
    ' Open Browser and navigate to URL
    Call OpenBrowser(url) 
    ' Perform Login
    Call PerformLogin(username, password)
    ' Verify Login
    If VerifyLogin() Then
        ' Add Product to Cart and Checkout
        Call AddProductToCartAndCheckout(productName)
    End If
    ' Close Browser
    Call CloseBrowser()
End Sub
' Execute the test
Call TestCheckout()


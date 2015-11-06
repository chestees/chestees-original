<!--#include virtual="/includes/globalLib.asp" -->
<!--#include virtual="/includes/adovbs.inc" -->
<!--#include virtual="/includes/payPal.asp" -->
<%
Response.Buffer = False

OpenDB()

curPurchaseAmount = Request("PurchaseAmount")
curShippingCost = Request("ShippingCost")

strIPAddress = Request.ServerVariables("REMOTE_ADDR")
paymentType = "Sale"
strCurrencyCode = "USD"

strCouponCode = Request("CouponCode")
If strCouponCode <> "" Then

	SQL = "SELECT Discount, DollarAmount FROM tblCouponCode WHERE CouponCode = " & SQLEncode(strCouponCode)
		Set rsCoupon = Conn.Execute(SQL)
	
	If Not rsCoupon.EOF Then
		intDiscount = rsCoupon("Discount")
		intDollarAmount = rsCoupon("DollarAmount")
		If intDiscount > 0 Then
			curDiscountAmount = curPurchaseAmount * intDiscount/100
		ElseIf intDollarAmount > 0 Then
			curDiscountAmount = intDollarAmount
		End If
	Else
		curDiscountAmount = 0
	End If
	rsCoupon.Close
	Set rsCoupon = Nothing
Else
	curDiscountAmount = 0
End If

curTotalAmount = formatNumber(Request("TotalAmount"), 2)

dtDate = Now()

'_____________________________________________________________________________________________
'GET CREDIT CARD INFORMATION AND BILLING ADDRESS
SQL = "SELECT FName, LName, Address, Address2, City, State, Zip, Email " & _
	"FROM tblBillingAddress " & _
	"WHERE " & varBuyer & "ID = " & cBuyerID & _
	" AND Lock = 0"
	Set rsBillingAddress = Conn.Execute(SQL)
	
SQL = "SELECT CardType, CardNumber, ExpMo, ExpYear, CVV " & _
	"FROM tblPayment " & _
	"WHERE " & varBuyer & "ID = " & cBuyerID & _
	" AND Lock = 0"
	Set rsPayment = Conn.Execute(SQL)

	strFName = rsBillingAddress("FName")
	strLName = rsBillingAddress("LName")
	strAddress = rsBillingAddress("Address")
	strAddress2 = rsBillingAddress("Address2")
	strCity = rsBillingAddress("City")
	strState = rsBillingAddress("State")
	strZip = rsBillingAddress("Zip")
	strEmail = rsBillingAddress("Email")
	
	strCardType = rsPayment("CardType")
	strCardNumber = rsPayment("CardNumber")
	strExpMo = rsPayment("ExpMo")
	strExpYear = rsPayment("ExpYear")
	strExp = strExpMo & "20" & strExpYear
	strCVV = rsPayment("CVV")
	
	rsBillingAddress.Close
	Set rsBillingAddress = Nothing
	
	rsPayment.Close
	Set rsPayment = Nothing

Call CloseDB()

'-----------------------------------------------------------------------------
' Construct the request string that will be sent to PayPal.
' The variable $nvpstr contains all the variables and is a
' name value pair string with &as a delimiter
'-----------------------------------------------------------------------------
nvpstr	=	"&PAYMENTACTION=" & paymentType & _
			"&AMT="& curTotalAmount &_
			"&CREDITCARDTYPE="& strCardType &_
			"&ACCT="& strCardNumber & _
			"&EXPDATE=" & strExp &_
			"&CVV2=" & strCVV &_
			"&FIRSTNAME=" & strFName &_
			"&LASTNAME=" & strLName &_
			"&STREET=" & strAddress & strAddress2 &_
			"&CITY=" & strCity &_
			"&STATE=" & strState &_
			"&ZIP=" & strZip &_
			"&COUNTRYCODE=US" &_
			"&CURRENCYCODE=" & strCurrencyCode
nvpstr	=	URLEncode(nvpstr)

'response.Write("NVP: " & nvpstr)
'response.Flush()

'-----------------------------------------------------------------------------
' Make the API call to PayPal,using API signature.
' The API response is stored in an associative array called gv_resArray
'-----------------------------------------------------------------------------
Set resArray = hash_call("doDirectPayment",nvpstr)
ack = UCase(resArray("ACK"))

'-----------------------------------------------------------------------------
'FOR TESTING
'-----------------------------------------------------------------------------
'response.Write("NVP: " & nvpstr)
'response.Flush()
'response.Redirect("http://www.chestees.com/index.asp?ACK=" & ack)
'ack = "SUCCESS"
	
'----------------------------------------------------------------------------------
' Display the API request and API response back to the browser.
' If the response from PayPal was a success, display the response parameters
' If the response was an error, display the errors received
'----------------------------------------------------------------------------------
If ack="SUCCESS" Then		
	'_____________________________________________________________________________________________
	'OPEN DATABASE CONNECTION
	Call OpenDB()
	
	'_____________________________________________________________________________________________
	'GET THE BOUGHT ITEMS FROM THE CART
	SQL = "SELECT C.ProductID, C.ProductStyleID, C.ProductSizeID, C.CartID, C.Quantity, B.BillingID, S.ShippingID, P.PaymentID " & _
		"FROM ((tblCart C INNER JOIN tblBillingAddress B ON C." & varBuyer & "ID = B." & varBuyer & "ID) " & _
		"INNER JOIN tblShippingAddress S ON C." & varBuyer & "ID = S." & varBuyer & "ID) " & _
		"INNER JOIN tblPayment P ON C." & varBuyer & "ID = P." & varBuyer & "ID " & _
		"WHERE C." & varBuyer & "ID = " & cBuyerID & " AND C.Purchased = 0 AND B.Lock = 0 AND S.Lock = 0 AND P.Lock = 0"
		Set rsCart = Conn.Execute(SQL)

		intBillingID = rsCart("BillingID")
		intShippingID = rsCart("ShippingID")
		intPaymentID = rsCart("PaymentID")
		
	'_____________________________________________________________________________________________
	'CREATE THE ORDER RECORD
	SQL = "INSERT INTO tblOrder (" & varBuyer & "ID, BillingID, ShippingID, PaymentID, PurchaseAmount, ShippingCost, DiscountAmount, TotalAmount, DateOrdered, Lock) VALUES (" & _
		SQLNumEncode(cBuyerID) & ", " & _
		SQLNumEncode(intBillingID) & ", " & _
		SQLNumEncode(intShippingID) & ", " & _
		SQLNumEncode(intPaymentID) & ", " & _
		SQLNumEncode(curPurchaseAmount) & ", " & _
		SQLNumEncode(curShippingCost) & ", " & _
		SQLNumEncode(curDiscountAmount) & ", " & _
		SQLNumEncode(curTotalAmount) & ", " & _
		SQLDateEncode(dtDate) & ", " & _
		1 & ")"
		Conn.Execute(SQL)

	'_____________________________________________________________________________________________
	'GETS THE ID OF THE ORDER JUST CREATED
	SQL = "SELECT Max(OrderID) AS MaxID FROM tblOrder"
		Set rsMaxID = Conn.Execute(SQL)
		intOrderID = rsMaxID("MaxID")
		
		Response.Cookies(cSiteName)("OrderID") = intOrderID
		
		rsMaxID.Close
		Set rsMaxID = Nothing	
	
	'_____________________________________________________________________________________________
	'UPDATE THE CART BY SETTING PURCHASED ITEMS TO "PURCHASED" AND THE RELATIONAL TABLE
	'UPDATE AVAILABLE QUANTITY
	Do While Not rsCart.EOF
		
		intCartID = rsCart("CartID")
		intProductID = rsCart("ProductID")
		intProductStyleID = rsCart("ProductStyleID")
		intProductSizeID = rsCart("ProductSizeID")
		intQuantity = rsCart("Quantity")
		
		SQL = "INSERT INTO relCartToOrder (CartID, OrderID) VALUES (" & _
			SQLNumEncode(intCartID) & ", " & _
			SQLNumEncode(intOrderID) & ")"
			Conn.Execute(SQL)
		
		'REDUCE THE QUANTITY	
		SQL = "UPDATE tblProductDetail SET Quantity = Quantity - " & intQuantity & " " & _
			"WHERE ProductID = " & intProductID & " " & _
			"AND ProductStyleID = " & intProductStyleID & " " & _
			"AND ProductSizeID = " & intProductSizeID
			Conn.Execute(SQL)
			
		SQL = "UPDATE tblCart SET Purchased = 1 WHERE CartID = " & intCartID
			Conn.Execute(SQL)
			
	rsCart.MoveNext
	Loop
	
	rsCart.Close
	Set rsCart = Nothing

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	'DELETE CREDIT CARD INFORMATION FROM THE DB!!!!
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	SQL = "SELECT CardNumber FROM tblPayment WHERE " & varBuyer & "ID = " & cBuyerID & _
		" AND Lock = 0"
		Set rsPayment = Conn.Execute(SQL)
		strCardNumber = right(rsPayment("CardNumber"),4)
		strCardNumber = "XXXX-XXXX-XXXX-" & strCardNumber
		rsPayment.Close
		Set rsPayment = Nothing
		
	SQL = "UPDATE tblPayment SET " & _
		"CardNumber = " & SQLEncode(strCardNumber) & ", " & _
		"CVV = '000', " & _
		"Lock = 1 " & _
		"WHERE " & varBuyer & "ID = " & cBuyerID & _
		" AND Lock = 0"
		Conn.Execute(SQL)
		
	SQL = "UPDATE tblBillingAddress SET Lock = 1 " & _
		"WHERE " & varBuyer & "ID = " & cBuyerID & _
		" AND Lock = 0"
		Conn.Execute(SQL)
		
	SQL = "UPDATE tblShippingAddress SET Lock = 1 " & _
		"WHERE " & varBuyer & "ID = " & cBuyerID & _
		" AND Lock = 0"
		Conn.Execute(SQL)
	
	CloseDB()
	
	Email_Confirm()
	
	'_____________________________________________________________________________________________
	'CLEAR THE VISITOR COOKIE
	If cVisitorID > 0 Then
		Response.Cookies(cSiteName)("VisitorID") = 0
	End If
	
	Response.Write("Success")
	
Else
	
	Set SESSION("nvpErrorResArray") = resArray
	'Response.Redirect "APIError.asp"
	 
	'--------------------------------------------------------------------------------------------
	' API Request and Response/Error Output
	' =====================================
	' This page will be called after getting Response from the server
	' or any Error occured during comminication for all APIs,to display Request,Response or Errors.
	'--------------------------------------------------------------------------------------------
	Dim resArray
	Dim message
	Dim ResponseHeader
	Dim Sepration
	message		 =  SESSION("msg")
	Sepration		=":"
	Set resArray = SESSION("nvpErrorResArray")
	
	ResponseHeader="Error Response Details"
	
	If Not SESSION("ErrorMessage")Then
		message = SESSION("ErrorMessage")
		ResponseHeader = ""
		Sepration= ""
	End If
	
	If Err.Number <> 0 Then
		SESSION("nvpReqArray") = Null
		
		Response.flush
	End If
	
	'reskey = resArray.Keys
	resitem = resArray.items
	
	strSQLError = ""

	For resindex = 0 To resArray.Count - 1
		
		strSQLError = strSQLError & ", " & resitem(resindex)
		If resindex = 7 Then
			strError = strError & resitem(resindex)
		End If
	Next

	strError = "ERROR(s): Credit card submitted was not accepted for the following reason:<br><span style='italic'>" & strError & "</span>"
	
	OpenDB()
	SQL = "INSERT INTO tblError (Items, VisitorID, CustomerID) VALUES (" & _
		SQLEncode(strSQLError) & ", " & SQLNumEncode(cVisitorID) & ", " & SQLNumEncode(cCustomerID) & ")"
		Conn.Execute(SQL)
	CloseDB()
	
	Response.write(strError)
	
End If
%>
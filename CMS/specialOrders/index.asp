<!--#include virtual="/includes/globalLib.asp"-->
<!--#include virtual="/includes/adovbs.inc" -->
<%
Call OpenDB()

'_____________________________________________________________________________________________
'Get Variables

btnSubmit = Request("Submit")

'_____________________________________________________________________________________________
'ADD Record

If btnSubmit <> "" Then

	varBuyer = "Visitor"
	cBuyerID = Request("VisitorID")
	'response.Write("V: " & cBuyerID)
	'response.Flush()
	cVisitorID = cBuyerID
	cCustomerID = 0
	
	strB_FName = trim(Request("B_FName"))
	strB_LName = trim(Request("B_LName"))
	strB_Email = trim(Request("B_Email"))
	strB_Email_Confirm = trim(Request("B_Email_Confirm"))
	strB_Address = trim(Request("B_Address"))
	strB_Address2 = trim(Request("B_Address2"))
	strB_City = trim(Request("B_City"))
	strB_State = trim(Request("B_State"))
	strB_Zip = trim(Request("B_Zip"))
	
	blnSameAsBilling = Request("SameAsBilling")
	
	strS_FName = trim(Request("S_FName"))
	strS_LName = trim(Request("S_LName"))
	strS_Address = Request(trim("S_Address"))
	strS_Address2 = trim(Request("S_Address2"))
	strS_City = trim(Request("S_City"))
	strS_State = trim(Request("S_State"))
	strS_Zip = trim(Request("S_Zip"))
	
	strCardType = Request("CardType")
	strCardNumber = trim(Request("CardNumber"))
	strExpMo = trim(Request("ExpMo"))
	strExpYear = trim(Request("ExpYear"))
	strCVV = trim(Request("CardCVV"))
	strCouponCode = trim(Request("CouponCode"))
%>
<!--
	'Billing Errors
	If strB_FName = "" Then
		ErrorMessage = ErrorMessage & "Error: Please provide your billing first name.<br>"
		intError = intError + 1
	End If
	If strB_LName = "" Then
		ErrorMessage = ErrorMessage & "Error: Please provide your billing last name.<br>"
		intError = intError + 1
	End If
	If strB_Email = "" Then
		ErrorMessage = ErrorMessage & "Error: Please provide your email address.<br>"
		intError = intError + 1
	End If
	If strB_Email <> strB_Email_Confirm Then
		ErrorMessage = ErrorMessage & "Error: The confirm email does not match.<br>"
		intError = intError + 1
	End If
	If strB_Address = "" Then
		ErrorMessage = ErrorMessage & "Error: Please provide your billing address.<br>"
		intError = intError + 1
	End If
	If strB_City = "" Then
		ErrorMessage = ErrorMessage & "Error: Please provide your billing city.<br>"
		intError = intError + 1
	End If
	If strB_State = "0" Then
		ErrorMessage = ErrorMessage & "Error: Please provide your billing state.<br>"
		intError = intError + 1
	End If
	If strB_Zip = "" Then
		ErrorMessage = ErrorMessage & "Error: Please provide your billing zip code.<br>"
		intError = intError + 1
	End If
	'Shipping Errors
	If strS_FName = "" AND blnSameAsBilling <> "ON" Then
		ErrorMessage = ErrorMessage & "Error: Please provide your shipping first name.<br>"
		intError = intError + 1
	End If
	If strS_LName = "" AND blnSameAsBilling <> "ON" Then
		ErrorMessage = ErrorMessage & "Error: Please provide your shipping last name.<br>"
		intError = intError + 1
	End If
	If strS_Address = "" AND blnSameAsBilling <> "ON" Then
		ErrorMessage = ErrorMessage & "Error: Please provide your shipping address.<br>"
		intError = intError + 1
	End If
	If strS_City = "" AND blnSameAsBilling <> "ON" Then
		ErrorMessage = ErrorMessage & "Error: Please provide your shipping city.<br>"
		intError = intError + 1
	End If
	If strS_State = "0" AND blnSameAsBilling <> "ON" Then
		ErrorMessage = ErrorMessage & "Error: Please provide your shipping state.<br>"
		intError = intError + 1
	End If
	If strS_Zip = "" AND blnSameAsBilling <> "ON" Then
		ErrorMessage = ErrorMessage & "Error: Please provide your shipping zip code.<br>"
		intError = intError + 1
	End If
	If strCardType = "0" OR strCardNumber = "" OR strExpMo = "0" OR strExpYear = "0" OR strCVV = "" Then
		ErrorMessage = ErrorMessage & "Error: Some credit card information was left blank. Please fill in all fields.<br>"
		intError = intError + 1
	End If
-->
<%
	'_____________________________________________________________________________________________
	'OPEN DATABASE CONNECTION
	Call OpenDB()
	
		Set cmd = Server.CreateObject("ADODB.Command")
		Set cmd.ActiveConnection = Conn
		cmd.CommandText = "usp_InsertBilling"
		cmd.CommandType = adCmdStoredProc
		
		cmd.Parameters.Append cmd.CreateParameter("VisitorID",adInteger,adParamInput)
		cmd.Parameters("VisitorID") = cVisitorID
		cmd.Parameters.Append cmd.CreateParameter("CustomerID",adInteger,adParamInput)
		cmd.Parameters("CustomerID") = cCustomerID
		cmd.Parameters.Append cmd.CreateParameter("FName",adVarChar,adParamInput,50)
		cmd.Parameters("FName") = strB_FName
		cmd.Parameters.Append cmd.CreateParameter("LName",adVarChar,adParamInput,50)
		cmd.Parameters("LName") = strB_LName
		cmd.Parameters.Append cmd.CreateParameter("Email",adVarChar,adParamInput,100)
		cmd.Parameters("Email") = strB_Email
		cmd.Parameters.Append cmd.CreateParameter("Address",adVarChar,adParamInput,125)
		cmd.Parameters("Address") = strB_Address
		cmd.Parameters.Append cmd.CreateParameter("Address2",adVarChar,adParamInput,125)
		cmd.Parameters("Address2") = strB_Address2
		cmd.Parameters.Append cmd.CreateParameter("City",adVarChar,adParamInput,75)
		cmd.Parameters("City") = strB_City
		cmd.Parameters.Append cmd.CreateParameter("State",adVarChar,adParamInput,2)
		cmd.Parameters("State") = strB_State
		cmd.Parameters.Append cmd.CreateParameter("Zip",adVarChar,adParamInput,5)
		cmd.Parameters("Zip") = strB_Zip
	
		cmd.Execute
		set cmd = nothing

		If blnSameAsBilling = "ON" Then
				
			Set cmd = Server.CreateObject("ADODB.Command")
			Set cmd.ActiveConnection = Conn
			cmd.CommandText = "usp_InsertShipping"
			cmd.CommandType = adCmdStoredProc
			
			cmd.Parameters.Append cmd.CreateParameter("VisitorID",adInteger,adParamInput)
			cmd.Parameters("VisitorID") = cVisitorID
			cmd.Parameters.Append cmd.CreateParameter("CustomerID",adInteger,adParamInput)
			cmd.Parameters("CustomerID") = cCustomerID
			cmd.Parameters.Append cmd.CreateParameter("FName",adVarChar,adParamInput,50)
			cmd.Parameters("FName") = strB_FName
			cmd.Parameters.Append cmd.CreateParameter("LName",adVarChar,adParamInput,50)
			cmd.Parameters("LName") = strB_LName
			cmd.Parameters.Append cmd.CreateParameter("Address",adVarChar,adParamInput,125)
			cmd.Parameters("Address") = strB_Address
			cmd.Parameters.Append cmd.CreateParameter("Address2",adVarChar,adParamInput,125)
			cmd.Parameters("Address2") = strB_Address2
			cmd.Parameters.Append cmd.CreateParameter("City",adVarChar,adParamInput,75)
			cmd.Parameters("City") = strB_City
			cmd.Parameters.Append cmd.CreateParameter("State",adVarChar,adParamInput,2)
			cmd.Parameters("State") = strB_State
			cmd.Parameters.Append cmd.CreateParameter("Zip",adVarChar,adParamInput,5)
			cmd.Parameters("Zip") = strB_Zip
		
			cmd.Execute
			set cmd = nothing
		
		Else
		
			Set cmd = Server.CreateObject("ADODB.Command")
			Set cmd.ActiveConnection = Conn
			cmd.CommandText = "usp_InsertShipping"
			cmd.CommandType = adCmdStoredProc
			
			cmd.Parameters.Append cmd.CreateParameter("VisitorID",adInteger,adParamInput)
			cmd.Parameters("VisitorID") = cVisitorID
			cmd.Parameters.Append cmd.CreateParameter("CustomerID",adInteger,adParamInput)
			cmd.Parameters("CustomerID") = cCustomerID
			cmd.Parameters.Append cmd.CreateParameter("FName",adVarChar,adParamInput,50)
			cmd.Parameters("FName") = strS_FName
			cmd.Parameters.Append cmd.CreateParameter("LName",adVarChar,adParamInput,50)
			cmd.Parameters("LName") = strS_LName
			cmd.Parameters.Append cmd.CreateParameter("Address",adVarChar,adParamInput,125)
			cmd.Parameters("Address") = strS_Address
			cmd.Parameters.Append cmd.CreateParameter("Address2",adVarChar,adParamInput,125)
			cmd.Parameters("Address2") = strS_Address2
			cmd.Parameters.Append cmd.CreateParameter("City",adVarChar,adParamInput,75)
			cmd.Parameters("City") = strS_City
			cmd.Parameters.Append cmd.CreateParameter("State",adVarChar,adParamInput,2)
			cmd.Parameters("State") = strS_State
			cmd.Parameters.Append cmd.CreateParameter("Zip",adVarChar,adParamInput,5)
			cmd.Parameters("Zip") = strS_Zip
		
			cmd.Execute
			set cmd = nothing
		
		End If
	
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = Conn
	cmd.CommandText = "usp_InsertPayment"
	cmd.CommandType = adCmdStoredProc
	
	cmd.Parameters.Append cmd.CreateParameter("VisitorID",adInteger,adParamInput)
	cmd.Parameters("VisitorID") = cVisitorID
	cmd.Parameters.Append cmd.CreateParameter("CustomerID",adInteger,adParamInput)
	cmd.Parameters("CustomerID") = cCustomerID
	cmd.Parameters.Append cmd.CreateParameter("CardType",adVarChar,adParamInput,50)
	cmd.Parameters("CardType") = strCardType
	cmd.Parameters.Append cmd.CreateParameter("CardNumber",adVarChar,adParamInput,50)
	cmd.Parameters("CardNumber") = strCardNumber
	cmd.Parameters.Append cmd.CreateParameter("ExpMo",adVarChar,adParamInput,2)
	cmd.Parameters("ExpMo") = strExpMo
	cmd.Parameters.Append cmd.CreateParameter("ExpYear",adVarChar,adParamInput,4)
	cmd.Parameters("ExpYear") = strExpYear
	cmd.Parameters.Append cmd.CreateParameter("CVV",adVarChar,adParamInput,4)
	cmd.Parameters("CVV") = strCVV
	cmd.Parameters.Append cmd.CreateParameter("CouponCode",adVarChar,adParamInput,50)
	cmd.Parameters("CouponCode") = Ucase(strCouponCode)
			
	cmd.Execute
	set cmd = nothing
		
	
	'_____________________________________________________________________________________________
	'GET THE BOUGHT ITEMS FROM THE CART
	SQL = "SELECT C.ProductID, C.ProductStyleID, C.ProductSizeID, C.CartID, C.Quantity, B.BillingID, S.ShippingID, P.PaymentID " & _
		"FROM ((tblCart C INNER JOIN tblBillingAddress B ON C." & varBuyer & "ID = B." & varBuyer & "ID) " & _
		"INNER JOIN tblShippingAddress S ON C." & varBuyer & "ID = S." & varBuyer & "ID) " & _
		"INNER JOIN tblPayment P ON C." & varBuyer & "ID = P." & varBuyer & "ID " & _
		"WHERE C." & varBuyer & "ID = " & cBuyerID & " AND C.Purchased = 0 AND B.Lock = 0 AND S.Lock = 0 AND P.Lock = 0"
		Set rsCart = Conn.Execute(SQL)

		'response.Write(SQL)
		'response.Flush()
		
		intBillingID = rsCart("BillingID")
		intShippingID = rsCart("ShippingID")
		intPaymentID = rsCart("PaymentID")
	
	curPurchaseAmount = Request("PurchaseAmount")
	curShippingCost = Request("ShippingAmount")
	curDiscountAmount = Request("DiscountAmount")	
	curTotalAmount = formatNumber(Request("TotalAmount"), 2)

	dtDate = DateAdd("h", +3, Now())
		
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
			
		SQL = "UPDATE tblCart SET Purchased = 1 WHERE CartID = " & intCartID
			Conn.Execute(SQL)
			
	rsCart.MoveNext
	Loop
	
	rsCart.Close
	Set rsCart = Nothing

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	'DELETE CREDIT CARD INFORMATION FROM THE DB!!!!
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	strCardNumber = "XXXX-XXXX-XXXX-0000"

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

	If intError = 0 Then
	
		Response.Redirect("../orders/")

	Else
		
		response.Write("ERROR: "+ ErrorMessage)
		Response.Flush()
		'_____________________________________________________________________________________________
		'OPEN DATABASE CONNECTION
		Call OpenDB()
	
		SQL = "SELECT FName, LName, Address, Address2, City, State, Zip, Email FROM tblBillingAddress WHERE Lock = 0 AND " & varBuyer & "ID = " & cBuyerID
			Set rsBilling = Conn.Execute(SQL)
			
			If Not rsBilling.EOF Then
			strB_FName = rsBilling("FName")
			strB_LName = rsBilling("LName")
			strB_Address = rsBilling("Address")
			strB_Address2 = rsBilling("Address2")
			strB_City = rsBilling("City")
			strB_State = rsBilling("State")
			strB_Zip = rsBilling("Zip")
			strB_Email = rsBilling("Email")
			End If
			rsBilling.Close
			Set rsBilling = Nothing
			
		SQL = "SELECT FName, LName, Address, Address2, City, State, Zip FROM tblShippingAddress WHERE Lock = 0 AND " & varBuyer & "ID = " & cBuyerID
			Set rsShipping = Conn.Execute(SQL)
			
			If Not rsShipping.EOF Then
			strS_FName = rsShipping("FName")
			strS_LName = rsShipping("LName")
			strS_Address = rsShipping("Address")
			strS_Address2 = rsShipping("Address2")
			strS_City = rsShipping("City")
			strS_State = rsShipping("State")
			strS_Zip = rsShipping("Zip")
			End If
			rsShipping.Close
			Set rsShipping = Nothing
	
		'SQL = "DELETE FROM tblPayment WHERE Lock = 0 AND " & varBuyer & "ID = " & cBuyerID
		'	Conn.Execute(SQL)
		
		CloseDB()
		
	End If

End If
%>
<html>
<head>
<title><%=cFriendlySiteName%> | Administration</title>
<link rel="stylesheet" href="/css/stylesheet.css">
<link rel="shortcut icon" href="/images/favicon.ico" type="image/x-icon">
</head>

<body leftmargin="0" topmargin="0" marginWidth="0" marginHeight="0">
<table width="804" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
  	<td bgcolor="#A13846" width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
	<td bgcolor="#EEF2FC">
	  <table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
	  	<tr>
		  <td colspan="3"><img src="<%=cAdminPath%>images/header.jpg"></td>
		</tr>
		<tr>
		  <td width="153" valign="top">
		  	<table border="0" cellspacing="0" cellpadding="0">
			  <tr>
			  	<td align="center">
				  <table border="0" cellspacing="0" cellpadding="0">
				  	<tr>
				  	  <td width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="38"></td>
					  <td width="151" class="NavTitle">Navigation</td>
					</tr>
				  </table>
				</td>
			  </tr>
			  <tr>
			  	<td>
<!--#include virtual="/incNav.asp" -->
				</td>
			  </tr>
			</table>
		  </td>
		  <td width="2" bgcolor="#A13846"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
		  <td width="645" valign="top">
		  	<table border="0" cellspacing="0" cellpadding="0" width="645">
			  <tr>
			  	<td>
				  <table border="0" cellspacing="0" cellpadding="0">
				  	<tr>
				  	  <td width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="38"></td>
					  <td width="621" class="PageTitle" align="right">
					  	SPECIAL ORDERS :: ADD
					  </td>
					  <td width="20" bgcolor="#A13846"><img src="<%=cAdminPath%>images/filler.gif" width="20" height="8"></td>
					  <td width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td colspan="4"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="20"></td>
					</tr>
					<tr>
					  <td><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
					  <td colspan="2" bgcolor="#A13846"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="2"></td>
					  <td><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
					  <td colspan="2" align="center" bgcolor="#E4DAC4">~ <a class="Nav" href="index.asp">Back to listing</a> ~</td>
					  <td><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
					  <td colspan="2" bgcolor="#A13846"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="2"></td>
					  <td><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td colspan="4"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="10"></td>
					</tr>
				  </table>
				</td>
			  </tr>
			  <tr>
			  	<td>
				  <table border="0" cellspacing="0" cellpadding="4" width="599" align="center">
<form action="index.asp" method="post">
		<tr>
          <td valign="top" align="center">	  	  
		  <table width="599" border="0" cellspacing="0" cellpadding="0" align="center">
<%
If intError > 0 Then
	Response.Write "<tr><td colspan=3 height=15><img src=images/filler.gif height=15></td></tr>"
  	Response.Write "<tr><td class=ErrorMessage colspan=3>" & ErrorMessage & "</td></tr>"
End If
%>
            <tr>
			  	<td colspan="2">
				  <table border="0" cellspacing="0" cellpadding="4" width="599" align="center">
					<tr bgcolor="#E2E6EF">
					  <td align="center">Cart ID: <input name="VisitorID" type="text" class="formField" style="width:250px" value=""></td>
					</tr>

					<tr>
					  <td colspan="4" style="border-top:1px dashed #A13846;"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="10"></td>
					</tr>
				  </table>
				</td>
			  </tr>
            <tr>
              <td valign="top">
              	<table border="0" cellspacing="0" cellpadding="0">
                  <tr>
				  <tr>
                    <td style="background-color:#CCCCCC; padding-top:8px; padding-bottom:8px; border-top:1px #C84848 solid; border-bottom:1px #C84848 solid;">&nbsp;<strong>Billing address</strong></td>
                  </tr>
                  <tr>
                    <td style="padding-top:5px;">
                      <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td>First name*<br>
                              <input name="B_FName" type="text" class="formField" style="width:150px" value="<%=strB_FName%>"></td>
                          <td width="10"><img src="/images/filler.gif" width="10" height="1"></td>
                          <td>Last name*<br>
                              <input name="B_LName" type="text" class="formField" style="width:150px" value="<%=strB_LName%>"></td>
                        </tr>
                    </table></td>
                  </tr>
				  <tr>
                    <td style="padding-top:10px;">Email*<br>
                        <input name="B_Email" type="text" class="formField" style="width:250px" value="<%=strB_Email%>"></td>
                  </tr>
				  <tr>
                    <td style="padding-top:10px;">Confirm Email*<br>
                        <input name="B_Email_Confirm" type="text" class="formField" style="width:250px" value="<%=strB_Email%>"></td>
                  </tr>
                  <tr>
                    <td style="padding-top:10px;">Address*<br>
                        <input name="B_Address" type="text" class="formField" style="width:250px" value="<%=strB_Address%>"></td>
                  </tr>
                  <tr>
                    <td style="padding-top:10px;">Address 2<br>
                        <input name="B_Address2" type="text" class="formField" style="width:250px" value="<%=strB_Address2%>"></td>
                  </tr>
                  <tr>
                    <td style="padding-top:10px;">City*<br>
                        <input name="B_City" type="text" class="formField" style="width:150px" value="<%=strB_City%>"></td>
                  </tr>
                  <tr>
                    <td style="padding-top:10px;">
                      <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td>State*<br>
                              <select name="B_State" class="formSelect" style="width:150px" size="1">
                                <option value="0">________</option>
                                <option value="UK">United Kingdom</option>
                                <option value="CD">Canada</option>
                                <option value="AU">Australia</option>
                                <option value="SW">Sweden</option>
<%
OpenDB()

SQL = "SELECT State, Abbrev FROM tblState"
	Set rsStates = Conn.Execute(SQL)

	Do While Not rsStates.EOF
	
		strState = rsStates("State")
		strAbbrev = rsStates("Abbrev")
		
		Response.Write("<option value=" & strAbbrev)
		If strB_State = strAbbrev Then
			Response.Write(" selected")
		End If
		Response.Write(">" & strState & "</option>" & vbCRlf)
	
	rsStates.MoveNext
	Loop
	
	rsStates.Close
	Set rsStates = Nothing
	
CloseDB()
%>
                            </select></td>
                          <td style="padding-left:10px;">Zip code*<br>
                              <input name="B_Zip" type="text" maxlength="5" class="formField" style="width:50px" value="<%=strB_Zip%>"></td>
                        </tr>
                    </table></td>
                  </tr>
              </table>
			  <table border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="8"><img src="/images/filler.gif" width="1" height="8"></td>
                  </tr>
				  <tr>
                    <td style="background-color:#CCCCCC; padding-top:8px; padding-bottom:8px; border-top:1px #C84848 solid; border-bottom:1px #C84848 solid;">&nbsp;<strong>Credit card information</strong></td>
                  </tr>
                  <tr>
                    <td style="padding-bottom:10px; padding-top:5px;">Type of Card*<br>
                      <select name="CardType" class="formSelect" style="width:150px" size="1">
                        <option value="0">________</option>
                        <option value="PayPal">PayPal</option>
                        <option value="Visa" selected>Visa</option>
                        <option value="MasterCard">MasterCard</option>
                        <option value="Discover">Discover</option>
                        <option value="Amex">American Express</option>
                        <option value="Cash">Cash</option>
                    </select></td>
                  </tr>
                  <tr>
                    <td style="padding-bottom:10px;">Card Number*<br>
                        <input name="CardNumber" type="text" class="formField" style="width:250px"></td>
                  </tr>
                  <tr>
                    <td style="padding-bottom:10px;">Expiration Date*<br>
                        MM <select name="ExpMo" size="1" class="formSelect" style="width:50px">
							<option value="0">___</option>
                            <option value="NA">NA</option>
							<option value="01">01</option>
							<option value="02">02</option>
							<option value="03">03</option>
							<option value="04">04</option>
							<option value="05">05</option>
							<option value="06">06</option>
							<option value="07">07</option>
							<option value="08">08</option>
							<option value="09">09</option>
							<option value="10">10</option>
							<option value="11">11</option>
							<option value="12">12</option>
							</select>
                        YYYY <select name="ExpYear" size="1" class="formSelect" style="width:80px">
							<option value="0">___</option>
                            <option value="NA">NA</option>
							<option value="08">08</option>
							<option value="09">09</option>
							<option value="10">10</option>
							<option value="11">11</option>
							<option value="12">12</option>
							<option value="13">13</option>
							<option value="14">14</option>
							<option value="15">15</option>
                            <option value="16">16</option>
                            <option value="17">17</option>
                            <option value="18">18</option>
							</select>
                    </td>
                  </tr>
                  <tr>
                    <td style="padding-bottom:10px;">Card Verification Number*<br>
                        <input name="CardCVV" type="text" class="formField" style="width:50px"></td>
                  </tr>
              </table>
			  </td>
              <td style="vertical-align:top; text-align:left; padding-left:10px;">
                <table border="0" cellspacing="0" cellpadding="0">
				  <tr>
                    <td style="background-color:#CCCCCC; padding-top:8px; padding-bottom:8px; border-top:1px #C84848 solid; border-bottom:1px #C84848 solid;">
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td>&nbsp;<b>Shipping address </b></td>
						  <td align="right"><input type="checkbox" name="SameAsBilling" value="ON"> Same as billing&nbsp;</td>
					 	 </tr>
						</table></td>
                  </tr>
                  <tr>
                    <td style="padding-top:5px;"><table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td>First name*<br>
                              <input name="S_FName" type="text" class="formField" style="width:150px" value="<%=strS_FName%>"></td>
                          <td style="padding-left:10px;">Last name*<br>
                              <input name="S_LName" type="text" class="formField" style="width:150px" value="<%=strS_LName%>"></td>
                        </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td style="padding-top:10px;">Address*<br>
                        <input name="S_Address" type="text" class="formField" style="width:250px" value="<%=strS_Address%>"></td>
                  </tr>
                  <tr>
                    <td style="padding-top:10px;">Address 2<br>
                        <input name="S_Address2" type="text" class="formField" style="width:250px" value="<%=strS_Address2%>"></td>
                  </tr>
                  <tr>
                    <td style="padding-top:10px;">City*<br>
                        <input name="S_City" type="text" class="formField" style="width:250px" value="<%=strS_City%>"></td>
                  </tr>
                  <tr>
                    <td style="padding-top:10px;"><table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td>State*<br>
                              <select name="S_State" class="formSelect" style="width:150px" size="1">
                                <option value="0">________</option>
                                <option value="UK">United Kingdom</option>
                                <option value="CD">Canada</option>
                                <option value="AU">Australia</option>
<%
OpenDB()

SQL = "SELECT State, Abbrev FROM tblState"
	Set rsStates = Conn.Execute(SQL)

	Do While Not rsStates.EOF
	
		strState = rsStates("State")
		strAbbrev = rsStates("Abbrev")
		
		Response.Write("<option value=" & strAbbrev)
		If strS_State = strAbbrev Then
			Response.Write(" selected")
		End If
		Response.Write(">" & strState & "</option>")
	
	rsStates.MoveNext
	Loop
	
	rsStates.Close
	Set rsStates = Nothing
	
CloseDB()
%>
                            </select></td>
                          <td style="padding-left:10px;">Zip code*<br>
                              <input name="S_Zip" type="text" maxlength="5" class="formField" style="width:50px" value="<%=strS_Zip%>"></td>
                        </tr>
                    </table></td>
                  </tr>
                </table>
			  	<table border="0" cellspacing="0" cellpadding="0" style="padding-top:10px;">
				  <tr>
                    <td style="background-color:#CCCCCC; padding-top:8px; padding-bottom:8px; border-top:1px #C84848 solid; border-bottom:1px #C84848 solid;">&nbsp;<strong>Discount code</strong></td>
                  </tr>
                  <tr>
                    <td style="padding-top:10px;">Coupon<br><input type="text" name="CouponCode" class="formField" style="width:150px"></td>
                  </tr>
                  <tr>
                    <td style="padding-top:10px;">Discount<br><input type="text" name="DiscountAmount" class="formField" style="width:150px" value="0"></td>
                  </tr>
                  <tr>
                    <td style="padding-top:10px;">Purchase Amount<br><input type="text" name="PurchaseAmount" class="formField" style="width:150px" value="0"></td>
                  </tr>
                  <tr>
                    <td style="padding-top:10px;">Shipping Amount<br><input type="text" name="ShippingAmount" class="formField" style="width:150px" value="0"></td>
                  </tr>
                  <tr>
                    <td style="padding-top:10px;">Total Amount<br><input type="text" name="TotalAmount" class="formField" style="width:150px" value="0"></td>
                  </tr>
				</table>
			  </td>
            </tr>
			<tr>
			  <td colspan="3" height="15"><img src="/images/filler.gif" width="1" height="15"></td>
			</tr>
			<tr>
			  <td colspan="3" height="1" bgcolor="#C84848"><img src="/images/filler.gif" width="1" height="1"></td>
			</tr>
			<tr>
			  <td colspan="3" style="text-align:right; padding-bottom:15px; padding-top:15px; border-top:1px #C84848 solid;">
                <input name="Submit" type="submit" class="formButton" style="width:125px" value="confirm >>"></td>
			</tr>
          </table></td>
   	    </tr>
</form>
				  </table>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
	<td bgcolor="#A13846" width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
  </tr>
  <tr>
  	<td colspan="3" bgcolor="#A13846" height="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="2"></td>
  </tr>
</table>
<!--#include virtual="/incFooter.asp" -->

</body>
</html>
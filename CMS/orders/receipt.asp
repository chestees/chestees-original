<!--#include virtual="/includes/globalLib.asp"-->
<%
intOrderID = Request("OrderID")
'Response.Write("COOKIE: " & Request.Cookies(cSiteName)("OrderID"))
%>
<html>
<head>
<title><%=cFriendlySiteName%> | Administration</title>
<link rel="stylesheet" href="/css/stylesheet.css">
<link rel="shortcut icon" href="/images/favicon.ico" type="image/x-icon">
</head>

<body leftmargin="0" topmargin="0" marginWidth="0" marginHeight="0">
<table width="582" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFFFFF">
  <tr>
  	<td width="1" bgcolor="#000000"><img src="http://www.chestees.com/images/filler.gif" width="1"></td>
  	<td width="580"><img src="http://www.chestees.com/images/header_PackingSlip.jpg"></td>
	<td width="1" bgcolor="#000000"><img src="http://www.chestees.com/images/filler.gif" width="1"></td>
  </tr>
  <tr>
  	<td width="1" bgcolor="#000000"><img src="http://www.chestees.com/images/filler.gif" width="1"></td>
	<td>    
	  <table cellpadding="0" cellspacing="0" border="0" width="580" align="center">
		<tr>
          <td valign="top" align="center">	  	  
		  <table width="550" border="0" cellspacing="0" cellpadding="0" align="center">
            <tr>
              <td width="550" valign="top">
			  	<table width="550" border="0" cellspacing="0" cellpadding="0">
					<tr>
					  <td colspan="4" height="1" bgcolor="#000000"><img src="http://www.chestees.com/images/filler.gif" width="1" height="1"></td>
					</tr>
					<tr>
					  <td style="padding-top:15px;">ORDER ID: <%=intOrderID%></td>
					</tr>
<%
Call OpenDB()

'_____________________________________________________________________________________________
'CREATE THE PRODUCTS RECORDSET
SQL = "SELECT C.CartID, P.Product, C.Price, S.SizeAbbr, S.ProductSizeID, Y.Style, C.Quantity, R.PurchaseAmount, R.ShippingCost, R.DiscountAmount, R.TotalAmount, R.DateOrdered "
	SQL = SQL & "FROM (((((tblCart C INNER JOIN tblProduct P ON C.ProductID = P.ProductID) "
	SQL = SQL & "INNER JOIN tblProductSize S ON C.ProductSizeID = S.ProductSizeID) "
	SQL = SQL & "INNER JOIN tblProductStyle Y ON C.ProductStyleID = Y.ProductStyleID)"
	SQL = SQL & "INNER JOIN relCartToOrder O ON C.CartID = O.CartID)"
	SQL = SQL & "INNER JOIN tblOrder R ON O.OrderID = R.OrderID) "
	SQL = SQL & "WHERE O.OrderID = " & intOrderID
	SQL = SQL & " ORDER BY S.ProductSizeID"
	Set rsCart = Conn.Execute(SQL)
	
If Not rsCart.EOF Then

	myColor = "#E2E2E2"
	
	curPurchaseAmount = rsCart("PurchaseAmount")
	curShippingCost = rsCart("ShippingCost")
	intDiscountAmount = rsCart("DiscountAmount")
	curTotalAmount = rsCart("TotalAmount")
	dtDateOrdered = rsCart("DateOrdered")
%>
					<tr>
					  <td style="padding-bottom:15px;">DATE ORDERED: <%=formatDateTime(dtDateOrdered,2)%></td>
					</tr>
					<tr bgcolor="#CCCCCC">
					  <td width="370"><b>Product</b></td>
					  <td width="60" align="center"><b>Size</b></td>
					  <td width="60" align="center"><b>Price</b></td>
					  <td width="60" align="center"><b>Qty</b></td>
					</tr>
					<tr>
					  <td colspan="4" height="1" bgcolor="#000000"><img src="http://www.chestees.com/images/filler.gif" height="1"></td>
					</tr>
					<tr>
					  <td colspan="4" height="15"><img src="http://www.chestees.com/images/filler.gif" width="1" height="15"></td>
					</tr>
<%
	Do While Not rsCart.EOF
	
		intCartID = rsCart("CartID")
		strProduct = rsCart("Product")
		strStyle = rsCart("Style")
		strSizeAbbr = rsCart("SizeAbbr")
		curPrice = rsCart("Price")
		intQuantity = rsCart("Quantity")
		
		intTotalQuantity = intTotalQuantity + intQuantity

		If myColor = "#E2E2E2" Then myColor = "#FFFFFF" Else myColor = "#E2E2E2"
%>
					<tr bgcolor="<%=myColor%>">
					  <td><b><%=strProduct%></b><br>
					    <span class="Black11px"><%=strStyle%></span></td>
					  <td align="center" valign="top"><%=strSizeAbbr%></td>
					  <td align="center" valign="top"><%=formatCurrency(curPrice,0)%></td>
					  <td align="center" valign="top"><%=intQuantity%></td>
					</tr>
<%
	rsCart.MoveNext
	Loop
	
End If
%>
					<tr>
					  <td colspan="4" height="1" bgcolor="#000000"><img src="http://www.chestees.com/images/filler.gif" height="1"></td>
					</tr>
				  </table>
			  	<table width="550" border="0" cellspacing="0" cellpadding="0">
<%
If cCustomerID > 0 Then
	varBuyer = "Customer"
	cBuyerID = cCustomerID
ElseIf cVisitorID > 0 Then
	varBuyer = "Visitor"
	cBuyerID = cVisitorID
End If
	
SQL = "SELECT B.FName, B.LName, B.Address, B.Address2, B.City, B.State, B.Zip, B.Email FROM tblBillingAddress B INNER JOIN tblOrder O ON O.BillingID = B.BillingID WHERE O.OrderID = " & intOrderID
	Set rsBilling = Conn.Execute(SQL)
	
	strB_FName = rsBilling("FName")
	strB_LName = rsBilling("LName")
	strB_Address = rsBilling("Address")
	strB_Address2 = rsBilling("Address2")
	strB_City = rsBilling("City")
	strB_State = rsBilling("State")
	strB_Zip = rsBilling("Zip")
	strB_Email = rsBilling("Email")
	
	rsBilling.Close
	Set rsBilling = Nothing
	
SQL = "SELECT S.FName, S.LName, S.Address, S.Address2, S.City, S.State, S.Zip FROM tblShippingAddress S INNER JOIN tblOrder O ON O.ShippingID = S.ShippingID WHERE O.OrderID = " & intOrderID
	Set rsShipping = Conn.Execute(SQL)
	
	strS_FName = rsShipping("FName")
	strS_LName = rsShipping("LName")
	strS_Address = rsShipping("Address")
	strS_Address2 = rsShipping("Address2")
	strS_City = rsShipping("City")
	strS_State = rsShipping("State")
	strS_Zip = rsShipping("Zip")
	
	rsShipping.Close
	Set rsShipping = Nothing
%>
                  <tr>
                    <td style="padding-top:10px; padding-bottom:8px;"><strong>billing address</strong></td>
                  </tr>
                  <tr>
                    <td style="border-bottom:1px #000 solid; padding-bottom:8px;"><%=strB_FName%>&nbsp;<%=strB_LName%><br><%=strB_Address%><%If strB_Address2 <> "" Then Response.Write("<br>" & strB_Address2)%><br><%=strB_City%>, <%=strB_State%>&nbsp;&nbsp;<%=strB_Zip%></td>
                  </tr>
				  <tr>
                    <td style="padding-top:10px; padding-bottom:8px;"><strong>shipping address</strong></td>
                  </tr>
                  <tr>
                    <td style="border-bottom:1px #000 solid; padding-bottom:10px;"><%=strS_FName%>&nbsp;<%=strS_LName%><br><%=strS_Address%><%If strS_Address2 <> "" Then Response.Write("<br>" & strS_Address2)%><br><%=strS_City%>, <%=strS_State%>&nbsp;&nbsp;<%=strS_Zip%></td>
                  </tr>
				  <tr>
                    <td style="padding-top:10px; padding-bottom:8px;"><strong>payment</strong></td>
                  </tr>
<%
SQL = "SELECT P.CardType, P.CardNumber, P.ExpMo, P.ExpYear, P.CouponCode FROM tblPayment P INNER JOIN tblOrder O ON O.PaymentID = P.PaymentID WHERE O.OrderID = " & intOrderID
	Set rsPayment = Conn.Execute(SQL)
	
	strCardType = rsPayment("CardType")
	strCardNumber = rsPayment("CardNumber")
	strCardNumber = right(strCardNumber,4)
	strCardNumber = "XXXX-XXXX-XXXX-" & strCardNumber
	strExpMo = rsPayment("ExpMo")
	strExpYear = rsPayment("ExpYear")
	strCouponCode = rsPayment("CouponCode")
	
	rsPayment.Close
	Set rsPayment = Nothing	
%>
                  <tr>
                    <td style="padding-bottom:15px; border-bottom:1px #000 solid;"><%If strCardType <> "PayPal" Then%><%=strB_FName%>&nbsp;<%=strB_LName%><br><%End If%>
                      <%=strCardType%>
                      <%If strCardType <> "PayPal" Then%><br>
                      <%=strCardNumber%><br>
                      <%=strExpMo%>/<%=strExpYear%>
                      <%End If%></td>
                  </tr>
              	</table>
			    <table width="550" border="0" cellspacing="0" cellpadding="0">
<%If strCouponCode <> "" Then

	SQL = "SELECT Discount FROM tblCouponCode WHERE CouponCode = " & SQLEncode(strCouponCode)
		Set rsCoupon = Conn.Execute(SQL)
	If Not rsCoupon.EOF Then
		intDiscount = rsCoupon("Discount")
		intDollarAmount = rsCoupon("DollarAmount")
		If intDiscount > 0 Then
			curDiscountAmount = curPurchaseAmount * intDiscount/100
			strDiscount = intDiscount & "% OFF"
		ElseIf intDollarAmount > 0 Then
			curDiscountAmount = intDollarAmount
			strDiscount = "$" & curDiscountAmount & " OFF"
		End If
		
		'curSubTotal = curTotal - curDiscountAmount
	Else
		'curSubTotal = curTotal
	End If
	rsCoupon.Close
	Set rsCoupon = Nothing
Else
	curSubTotal = curTotal
End If

Call CloseDB()
%>
					
					<tr bgcolor="#CCCCCC" style="padding:5px;">
					  <td colspan="5" valign="top" align="right" style="border-bottom:1px #000 solid;">
					  	<table border="0" cellspacing="0" cellpadding="0">
						  <tr>
						  	<td style="text-align:right;">Sub Total: </td>
							<td  style="text-align:right; width:100px;"><%=formatCurrency(curPurchaseAmount,2)%></td>
						  </tr>
						  <%If strCouponCode <> "" Then%>
						  <tr>
						  	<td style="text-align:right;">Discount (<%=strCouponCode%> = <%=strDiscount%>)</td>
							<td style="text-align:right;"><%=formatCurrency(curDiscountAmount,2)%></td>
						  </tr>
						  <%End If%>
						  <tr>
						  	<td style="padding-bottom:12px; text-align:right;">Shipping: </td>
						  	<td style="padding-bottom:12px; text-align:right;"><%=formatCurrency(curShippingCost,2)%></td>
						  </tr>
						  <tr style="font-weight:bold;">
						  	<td style="text-align:right;">Total amount: </td>
							<td style="text-align:right;"><%=formatCurrency(curTotalAmount)%></td>
						  </tr>
						</table>
					  </td>
					</tr>
				  </table>
			  </td>
            </tr>
			<tr>
			  <td style="padding-top:15px; padding-bottom:15px;">Thank you for your order! Since we want everyone to love their Chestees tee, if for some reason there is an issue with it, please use the contact form on our website within 30 days and explain the situation. Chestees will then facilitate an exchange or refund.</td>
			</tr>
          </table>
		</td>
   	  </tr>
    </table>
	</td>
	<td width="1" bgcolor="#000000"><img src="http://www.chestees.com/images/filler.gif" width="1"></td>
  </tr>
  <tr>
	<td colspan="3" height="1" bgcolor="#000000"><img src="http://www.chestees.com/images/filler.gif" height="1"></td>
  </tr>
</table>

</body>
</html>
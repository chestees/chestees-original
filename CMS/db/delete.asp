<!--#include virtual="/includes/globalLib.asp"-->
<%
'_____________________________________________________________________________________________
'Get Variables
btnSubmit = Request("Submit")

'_____________________________________________________________________________________________
'ADD Record
OpenDB()
If request("action") = 4 AND btnSubmit <> "" Then

	SQL = "delete from tblOrder WHERE OrderID = " & request("OrderID")
		Conn.Execute(SQL)
		'response.Write(SQL & "<br /><br />")
		
	If request("BillingID") <> "" Then
	SQL = "delete from tblBillingAddress WHERE BillingID = " & request("BillingID")
		Conn.Execute(SQL)
		'response.Write(SQL & "<br /><br />")
	End If
	If request("ShippingID") <> "" Then
	SQL = "delete from tblShippingAddress WHERE ShippingID = " & request("ShippingID")
		Conn.Execute(SQL)
		'response.Write(SQL & "<br /><br />")
	End If
	If request("PaymentID") <> "" Then
	SQL = "delete from tblPayment WHERE PaymentID = " & request("PaymentID")
		Conn.Execute(SQL)
		'response.Write(SQL & "<br /><br />")
	End If
	SQL = "delete from relCartToOrder WHERE OrderID = " & request("OrderID")
		Conn.Execute(SQL)
		'response.Write(SQL & "<br /><br />")
	response.Redirect("order.asp")
	
ElseIf request("action") = 1 AND btnSubmit <> "" Then

	SQL = "delete from tblBillingAddress WHERE BillingID = " & request("BillingID")
		Conn.Execute(SQL)
		'response.Write(SQL)
	response.Redirect("index.asp")
	
ElseIf request("action") = 2 AND btnSubmit <> "" Then

	SQL = "delete from tblShippingAddress WHERE ShippingID = " & request("ShippingID")
		Conn.Execute(SQL)
		'response.Write(SQL)
	response.Redirect("shipping.asp")
	
ElseIf request("action") = 3 AND btnSubmit <> "" Then

	SQL = "delete from tblPayment WHERE PaymentID = " & request("PaymentID")
		Conn.Execute(SQL)
		'response.Write(SQL)
	response.Redirect("payment.asp")
		
ElseIf request("action") = 5 AND btnSubmit <> "" Then
	
	SQL = "delete from relCartToOrder WHERE OrderID = " & request("OrderID")
		Conn.Execute(SQL)
		'response.Write(SQL)
	response.Redirect("relation.asp")
		
ElseIf request("action") = 6 AND btnSubmit <> "" Then
	
	SQL = "delete from tblCart WHERE CartID = " & request("CartID")
		Conn.Execute(SQL)
		'response.Write(SQL)
	response.Redirect("cart.asp")
		
End If
CloseDB()
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
					  <td width="621" class="PageTitle" align="right">DATABASE MANAGEMENT</td>
					  <td width="20" bgcolor="#A13846"><img src="<%=cAdminPath%>images/filler.gif" width="20" height="1"></td>
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
					  <td colspan="2" align="center" bgcolor="#E4DAC4"><a class="Nav" href="index.asp">Billing</a> ~ 
                      	<a class="Nav" href="shipping.asp">Shipping</a> ~ 
                        <a class="Nav" href="payment.asp">Payment</a> ~ 
                        <a class="Nav" href="order.asp">Order</a> ~ 
                        <a class="Nav" href="relation.asp">Relation</a></td>
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
				  	<tr bgcolor="#a13b3b">
					  <td colspan="2" class="White12pxB">DELETE RECORD</td>
					</tr>
<%
OpenDB()
If request("action") = 1 AND btnSubmit = "" Then

	SQL = "SELECT BillingID, CustomerID, VisitorID, FName, LName, Address, Address2, City, State, Zip, Email, Lock FROM tblBillingAddress WHERE BillingID = " & Request("BillingID")
		Set RS2 = Conn.Execute(SQL)
		
		Do While Not RS2.EOF
			
			intBillingID = RS2("BillingID")
			intCustomerID = RS2("CustomerID")
			intVisitorID = RS2("VisitorID")
			strB_FName = RS2("FName")
			strB_LName = RS2("LName")
			strB_Address = RS2("Address")
			strB_Address2 = RS2("Address2")
			strB_City = RS2("City")
			strB_State = RS2("State")
			strB_Zip = RS2("Zip")
			strB_Email = RS2("Email")
			blnLock = RS2("Lock")
			
			response.Write("<tr><td>Billing</td></tr>")
			response.Write("<tr><td valign=""top""><b>" & intBillingID & "</b>, C:" & intCustomerID & ", V:" & intVisitorID & ", " & strB_FName & ", " & strB_LName & ", " & strB_Address & ", " & strB_Address2 & ", " & strB_City & ", " & strB_State & ", " & strB_Zip & ", " & strB_Email & ", " & blnLock & "</td></tr>")
		
		RS2.MoveNext
		Loop
		
		RS2.Close
		Set RS2 = Nothing

ElseIf request("action") = 2 AND btnSubmit = "" Then

	SQL = "SELECT ShippingID, CustomerID, VisitorID, FName, LName, Address, Address2, City, State, Zip, Lock FROM tblShippingAddress WHERE ShippingID = " & Request("ShippingID")
		Set RS1 = Conn.Execute(SQL)
		
		If Not RS1.EOF Then
		Do While Not RS1.EOF
			
			intShippingID = RS1("ShippingID")
			intCustomerID = RS1("CustomerID")
			intVisitorID = RS1("VisitorID")
			strB_FName = RS1("FName")
			strB_LName = RS1("LName")
			strB_Address = RS1("Address")
			strB_Address2 = RS1("Address2")
			strB_City = RS1("City")
			strB_State = RS1("State")
			strB_Zip = RS1("Zip")
			blnLock = RS1("Lock")
			
			response.Write("<tr><td>Shipping</td></tr>")
			response.Write("<tr><td valign=""top""><b>" & intShippingID & "</b>, C:" & intCustomerID & ", V:" & intVisitorID & ", " & strB_FName & ", " & strB_LName & ", " & strB_Address & ", " & strB_Address2 & ", " & strB_City & ", " & strB_State & ", " & strB_Zip & ", " & blnLock & "</td></tr>")
		
		RS1.MoveNext
		Loop
		End If
		RS1.Close
		Set RS1 = Nothing
	
ElseIf request("action") = 3 AND btnSubmit = "" Then

	SQL = "SELECT PaymentID, CustomerID, VisitorID, CardType, CardNumber, ExpMo, ExpYear, CVV, CouponCode, Lock FROM tblPayment WHERE PaymentID = " & Request("PaymentID")
		Set RS3 = Conn.Execute(SQL)
		
		Do While Not RS3.EOF
			
			intPaymentID = RS3("PaymentID")
			intCustomerID = RS3("CustomerID")
			intVisitorID = RS3("VisitorID")
			strCardType = RS3("CardType")
			strCardNumber = RS3("CardNumber")
			strExpMo = RS3("ExpMo")
			strExpYear = RS3("ExpYear")
			strCVV = RS3("CVV")
			strCouponCode = RS3("CouponCode")
			blnLock = RS3("Lock")
			
			response.Write("<tr><td>Payment</td></tr>")
			response.Write("<tr><td valign=""top""><b>" & intPaymentID & "</b>, C:" & intCustomerID & ", V:" & intVisitorID & ", " & strCardType & ", " & strCardNumber & ", " & strExpMo & ", " & strExpYear & ", " & strCVV & ", " & strCouponCode & ", " & blnLock & "</td></tr>")
		
		RS3.MoveNext
		Loop
		
		RS3.Close
		Set RS3 = Nothing
		
ElseIf request("action") = 4 AND btnSubmit = "" Then
	
	SQL = "SELECT OrderID, CustomerID, VisitorID, ShippingID, BillingID, PaymentID, DateOrdered, Lock FROM tblOrder WHERE OrderID = " & Request("OrderID") 
		Set RS = Conn.Execute(SQL)
		
		If NOT RS.EOF Then
			
			intOrderID = RS("OrderID")
			intCustomerID = RS("CustomerID")
			intVisitorID = RS("VisitorID")
			intShippingID = RS("ShippingID")
			intBillingID = RS("BillingID")
			intPaymentID = RS("PaymentID")
			dtDateOrdered = RS("DateOrdered")
			blnLock = RS("Lock")
			
			response.Write("<tr><td>Order</td></tr>")
			response.Write("<tr><td valign=""top""><b>" & intOrderID & "</b>, C:" & intCustomerID & ", V:" & intVisitorID & ", B:" & intBillingID & ", S:" & intShippingID & ", P:" & intPaymentID & ", " & dtDateOrdered & ", " & blnLock & "</td></tr>")
		
		End If
		RS.Close
		Set RS = Nothing
	
	
	SQL = "SELECT ShippingID, CustomerID, VisitorID, FName, LName, Address, Address2, City, State, Zip, Lock FROM tblShippingAddress WHERE ShippingID = " & intShippingID
		Set RS1 = Conn.Execute(SQL)
		
		If Not RS1.EOF Then
		Do While Not RS1.EOF
			
			intShippingID = RS1("ShippingID")
			intCustomerID = RS1("CustomerID")
			intVisitorID = RS1("VisitorID")
			strB_FName = RS1("FName")
			strB_LName = RS1("LName")
			strB_Address = RS1("Address")
			strB_Address2 = RS1("Address2")
			strB_City = RS1("City")
			strB_State = RS1("State")
			strB_Zip = RS1("Zip")
			blnLock = RS1("Lock")
			
			response.Write("<tr><td>Shipping</td></tr>")
			response.Write("<tr><td valign=""top""><b>" & intShippingID & "</b>, C:" & intCustomerID & ", V:" & intVisitorID & ", " & strB_FName & ", " & strB_LName & ", " & strB_Address & ", " & strB_Address2 & ", " & strB_City & ", " & strB_State & ", " & strB_Zip & ", " & blnLock & "</td></tr>")
		
		RS1.MoveNext
		Loop
		End If
		RS1.Close
		Set RS1 = Nothing
	
	SQL = "SELECT BillingID, CustomerID, VisitorID, FName, LName, Address, Address2, City, State, Zip, Email, Lock FROM tblBillingAddress WHERE BillingID = " & intBillingID
		Set RS2 = Conn.Execute(SQL)
		
		Do While Not RS2.EOF
			
			intBillingID = RS2("BillingID")
			intCustomerID = RS2("CustomerID")
			intVisitorID = RS2("VisitorID")
			strB_FName = RS2("FName")
			strB_LName = RS2("LName")
			strB_Address = RS2("Address")
			strB_Address2 = RS2("Address2")
			strB_City = RS2("City")
			strB_State = RS2("State")
			strB_Zip = RS2("Zip")
			strB_Email = RS2("Email")
			blnLock = RS2("Lock")
			
			response.Write("<tr><td>Billing</td></tr>")
			response.Write("<tr><td valign=""top""><b>" & intBillingID & "</b>, C:" & intCustomerID & ", V:" & intVisitorID & ", " & strB_FName & ", " & strB_LName & ", " & strB_Address & ", " & strB_Address2 & ", " & strB_City & ", " & strB_State & ", " & strB_Zip & ", " & strB_Email & ", " & blnLock & "</td></tr>")
		
		RS2.MoveNext
		Loop
		
		RS2.Close
		Set RS2 = Nothing
	
	SQL = "SELECT PaymentID, CustomerID, VisitorID, CardType, CardNumber, ExpMo, ExpYear, CVV, CouponCode, Lock FROM tblPayment WHERE PaymentID = " & intPaymentID
		Set RS3 = Conn.Execute(SQL)
		
		Do While Not RS3.EOF
			
			intPaymentID = RS3("PaymentID")
			intCustomerID = RS3("CustomerID")
			intVisitorID = RS3("VisitorID")
			strCardType = RS3("CardType")
			strCardNumber = RS3("CardNumber")
			strExpMo = RS3("ExpMo")
			strExpYear = RS3("ExpYear")
			strCVV = RS3("CVV")
			strCouponCode = RS3("CouponCode")
			blnLock = RS3("Lock")
			
			response.Write("<tr><td>Payment</td></tr>")
			response.Write("<tr><td valign=""top""><b>" & intPaymentID & "</b>, C:" & intCustomerID & ", V:" & intVisitorID & ", " & strCardType & ", " & strCardNumber & ", " & strExpMo & ", " & strExpYear & ", " & strCVV & ", " & strCouponCode & ", " & blnLock & "</td></tr>")
		
		RS3.MoveNext
		Loop
		
		RS3.Close
		Set RS3 = Nothing
		
	SQL = "SELECT OrderID, CartID FROM relCartToOrder WHERE OrderID = " & request("OrderID")
		Set RS4 = Conn.Execute(SQL)
		
		Do While Not RS4.EOF
			
			intOrderID = RS4("OrderID")
			intCartID = RS4("CartID")
			
			response.Write("<tr><td>Relation</td></tr>")
			response.Write("<tr><td valign=""top""><b>OrderID:" & intOrderID & "</b>, CartID:" & intCartID & "</td></tr>")
		
		RS4.MoveNext
		Loop
		
		RS4.Close
		Set RS4 = Nothing

ElseIf request("action") = 5 AND btnSubmit = "" Then

	SQL = "SELECT OrderID, CartID FROM relCartToOrder WHERE OrderID = " & request("OrderID")
		Set RS4 = Conn.Execute(SQL)
		
		Do While Not RS4.EOF
			
			intOrderID = RS4("OrderID")
			intCartID = RS4("CartID")
			
			response.Write("<tr><td>Relation</td></tr>")
			response.Write("<tr><td valign=""top""><b>OrderID:" & intOrderID & "</b>, CartID:" & intCartID & "</td></tr>")
		
		RS4.MoveNext
		Loop
		
		RS4.Close
		Set RS4 = Nothing
	
End If
CloseDB()
%>
					<tr>
					  <td colspan="2"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="10"></td>
					</tr>
                    <tr><td>
<form method="post" action="delete.asp">
<input type="hidden" value="<%=request("OrderID")%>" name="OrderID" />
<input type="hidden" value="<%=request("PaymentID")%>" name="PaymentID" />
<input type="hidden" value="<%=request("BillingID")%>" name="BillingID" />
<input type="hidden" value="<%=request("ShippingID")%>" name="ShippingID" />
<input type="hidden" value="<%=request("CartID")%>" name="CartID" />
<input type="hidden" value="<%=request("action")%>" name="action" />
                    	<input type="submit" value="Submit" name="submit" />
</form>
                    </td></tr>
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
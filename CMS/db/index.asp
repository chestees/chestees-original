<!--#include virtual="/includes/globalLib.asp"-->
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
                        <a class="Nav" href="relation.asp">Relation</a> ~
                        <a class="Nav" href="cart.asp">Cart</a></td>
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
					  <td colspan="2" class="White12pxB">Billing</td>
					</tr>
<%
OpenDB()
SQL = "SELECT BillingID, CustomerID, VisitorID, FName, LName, Address, Address2, City, State, Zip, Email, Lock FROM tblBillingAddress ORDER BY BillingID DESC"
	Set RS = Conn.Execute(SQL)
	
	Do While Not RS.EOF
		
		intBillingID = RS("BillingID")
		intCustomerID = RS("CustomerID")
		intVisitorID = RS("VisitorID")
		strB_FName = RS("FName")
		strB_LName = RS("LName")
		strB_Address = RS("Address")
		strB_Address2 = RS("Address2")
		strB_City = RS("City")
		strB_State = RS("State")
		strB_Zip = RS("Zip")
		strB_Email = RS("Email")
		blnLock = RS("Lock")
		
		response.Write("<tr><td valign=""top""><a href=""delete.asp?action=1&BillingID="&intBillingID&""">X</a></td>")
		response.Write("<td valign=""top""><b>" & intBillingID & "</b>, C:" & intCustomerID & ", V:" & intVisitorID & ", " & strB_FName & ", " & strB_LName & ", " & strB_Address & ", " & strB_Address2 & ", " & strB_City & ", " & strB_State & ", " & strB_Zip & ", " & strB_Email & ", " & blnLock & "</td></tr>")
	
	RS.MoveNext
	Loop
	
	RS.Close
	Set RS = Nothing

Call CloseDB()
%>
					<tr>
					  <td colspan="2"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="10"></td>
					</tr>
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
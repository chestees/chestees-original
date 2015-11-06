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
					  <td width="621" class="PageTitle" align="right">ORDERS LISTING</td>
					  <td width="20" bgcolor="#A13846"><img src="<%=cAdminPath%>images/filler.gif" width="20" height="1"></td>
					  <td width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
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
<%
OpenDB()

SQL = "SELECT O.OrderID, O.CustomerID, O.DateOrdered, O.TotalAmount, O.ShippingCost, B.FName, B.LName FROM " & _
	"tblOrder O INNER JOIN tblBillingAddress B ON O.BillingID = B.BillingID ORDER BY O.DateOrdered DESC"
	Set	RS = Conn.Execute(SQL)
	
If Not RS.EOF Then
	
	Do While Not RS.EOF
	
	intOrderID = RS("OrderID")
	intCustomerID = RS("CustomerID")
	'intVisitorID = RS("VisitorID")
	dtDateOrdered = RS("DateOrdered")
	curTotalAmount = formatCurrency(RS("TotalAmount"),2)
	curShippingCost = formatCurrency(RS("ShippingCost"),2)
	strFName = RS("FName")
	strLName = RS("LName")
	
	If cColor = "#E2E6EF" Then cColor = "#EEF2FC" Else cColor = "#E2E6EF" End If
%>
					<tr bgcolor="<%=cColor%>">
					  <td><a class="Record" href="edit.asp?OrderID=<%=intOrderID%>"><%=strLName%>,&nbsp;<%=strFName%></a><br><%=dtDateOrdered%><br><b>Purchase Amount:</b> <%=curTotalAmount%></td>
					  <td width="150" align="right"><a href="print.asp?OrderID=<%=intOrderID%>">Print Receipt</a></td>
					</tr>
<%
	RS.MoveNext
	Loop
	
	RS.Close
	Set RS = nothing
	
Else
%>
					<tr bgcolor="#E2E6EF">
					  <td align="center">No Records</td>
					</tr>
<%
End If

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
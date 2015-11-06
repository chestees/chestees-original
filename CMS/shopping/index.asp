<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="/includes/globalLib.asp"-->
<%
intCartID = cInt(Request("CartID"))
blnDelete = Request("Delete")

If intCartID > 0 AND blnDelete = 1 Then

	OpenDB()
	
	SQL = "DELETE FROM tblCart WHERE CartID = " & intCartID
		Conn.Execute(SQL)
		
	CloseDB()
End If
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title><%=cFriendlySiteName%> | Administration</title>
<link rel="stylesheet" href="/css/stylesheet.css">
<link rel="shortcut icon" href="/images/favicon.ico" type="image/x-icon">
</head>

<body>
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
					  <td width="621" class="PageTitle" align="right">SHOPPING CART</td>
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

'_____________________________________________________________________________________________
'CREATE THE CART RECORDSET
SQL = "SELECT C.CustomerID, C.VisitorID, C.Purchased, C.DateAdded, C.CartID, P.Product, C.Price, S.SizeAbbr, Y.Style, C.Quantity "
	SQL = SQL & "FROM (((tblCart C INNER JOIN tblProduct P ON C.ProductID = P.ProductID) "
	SQL = SQL & "INNER JOIN tblProductSize S ON C.ProductSizeID = S.ProductSizeID) "
	SQL = SQL & "INNER JOIN tblProductStyle Y ON C.ProductStyleID = Y.ProductStyleID) "
	SQL = SQL & "ORDER BY C.DateAdded DESC"
	Set RS = Conn.Execute(SQL)

If Not RS.EOF Then
	i = 0
	Do While Not RS.EOF
	
	intCartID = RS("CartID")
	intCustomerID = RS("CustomerID")
	intVisitorID = RS("VisitorID")
	blnPurchased = RS("Purchased")
	dtDateAdded = RS("DateAdded")
	strProduct = RS("Product")
	strStyle = RS("Style")
	strSizeAbbr = RS("SizeAbbr")
	curPrice = RS("Price")
	intQuantity = RS("Quantity")
	
	i=i+1
	If cColor = "#E2E6EF" Then cColor = "#EEF2FC" Else cColor = "#E2E6EF" End If
%>
					<tr bgcolor="<%=cColor%>">
					  <td style="border-left:1px dashed #A13846;border-top:1px dashed #A13846;">
<%						If intCustomerID > 0 Then
							SQL = "SELECT Email FROM tblCustomer WHERE CustomerID = " & intCustomerID
								Set rsCustomer = Conn.Execute(SQL)
								
							response.Write("Customer: " & rsCustomer("Email"))
							rsCustomer.Close
							Set rsCustomer = Nothing
							 
						Else
							response.Write("Visitor: " & intVisitorID)
						End If
%>
<%If blnPurchased Then%>
					  <div class="Pink"><%=dtDateAdded%> - PURCHASED: <%=blnPurchased%></div>
<%Else%>
					  <div class="Black11px"><%=dtDateAdded%> - PURCHASED: <%=blnPurchased%></div>
<%End If%>
					  <b><%=strProduct%></b><br>
					    <span class="Black11px"><%=strStyle%></span>
<%If Not blnPurchased Then%>
                        <br /><a href="index.asp?CartID=<%=intCartID%>&Delete=1" class="Delete">delete</a>
<%End If%>
                      </td>
					  <td align="center" style="border-top:1px dashed #A13846; vertical-align:top;"><%=strSizeAbbr%></td>
					  <td align="center" style="border-top:1px dashed #A13846; vertical-align:top;"><%=formatCurrency(curPrice,0)%></td>
					  <td align="center" style="border-top:1px dashed #A13846; border-right:1px dashed #A13846; vertical-align:top;"><%=intQuantity%></td>
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
					  <td colspan="4" style="border-top:1px dashed #A13846;"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="10"></td>
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
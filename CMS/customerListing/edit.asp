<!--#include virtual="/includes/globalLib.asp"-->
<%
'_____________________________________________________________________________________________
'Get Variables

intCustomerID = Request("CustomerID")

If intCustomerID > 0 Then
	varBuyer = "Customer"
	cBuyerID = intCustomerID
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
					  	CUSTOMER :: INFO
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
					<tr bgcolor="#CCCCCC">
					  <td><b>Order #</b></td>
					  <td><b>Product</b></td>
					  <td><b>Size</b></td>
					  <td><b>Price</b></td>
					  <td><b>Shipping</b></td>
					  <td><b>Qty</b></td>
					  <td style="text-align:center"><b>Purchased</b></td>
					</tr>
<%
Call OpenDB()

'_____________________________________________________________________________________________
'CREATE THE PRODUCTS SIZES RECORDSET
SQL = "SELECT O.OrderID, C.CartID, P.Product, C.Price, S.Size, Y.Style, C.Quantity, C.Purchased FROM ((((tblCart C INNER JOIN tblProduct P ON C.ProductID = P.ProductID) "
	SQL = SQL & "INNER JOIN tblProductSize S ON C.ProductSizeID = S.ProductSizeID) "
	SQL = SQL & "INNER JOIN tblProductStyle Y ON C.ProductStyleID = Y.ProductStyleID)"
	SQL = SQL & "INNER JOIN relCartToOrder O ON C.CartID = O.CartID)"
	SQL = SQL & "WHERE C." & varBuyer & "ID = " & cBuyerID
	SQL = SQL & " ORDER BY C.Purchased DESC"
	Set rsCart = Conn.Execute(SQL)
	
If Not rsCart.EOF Then
	
	curTotal = 0
	curTotalShipping = 0
	intTotalQuantity = 0
	
	Do While Not rsCart.EOF
	
		intOrderID = rsCart("OrderID")
		intCartID = rsCart("CartID")
		strProduct = rsCart("Product")
		strSize = rsCart("Size")
		curPrice = rsCart("Price")
		intQuantity = rsCart("Quantity")
		blnPurchased = rsCart("Purchased")
		
		intTotalQuantity = intTotalQuantity + intQuantity
		curTotal = curTotal+curPrice*intQuantity
		If intQuantity > 3 Then
			cShippingCost = 0
		End If
%>
					<tr>
					  <td><a href="<%=cAdminPath%>orders/edit.asp?OrderID=<%=intOrderID%>"><%=intOrderID%></a></td>
					  <td><%=strProduct%></td>
					  <td><%=strSize%></td>
					  <td><%=formatCurrency(curPrice,0)%></td>
					  <td>&nbsp;</td>
					  <td><%=intQuantity%></td>
					  <td style="text-align:center"><%=blnPurchased%></td>					  
					</tr>
<%
	rsCart.MoveNext
	Loop
	
End If
Call CloseDB()
%>
					<tr>
					  <td colspan="5" height="10"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="10"></td>
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
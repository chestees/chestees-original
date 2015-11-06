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
  	<td bgcolor="#A13846" width="2"><img src="/images/filler.gif" width="2" height="1"></td>
	<td bgcolor="#EEF2FC">
	  <table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
	  	<tr>
		  <td colspan="3"><img src="/images/header.jpg"></td>
		</tr>
		<tr>
		  <td width="153" valign="top">
		  	<table border="0" cellspacing="0" cellpadding="0">
			  <tr>
			  	<td align="center">
				  <table border="0" cellspacing="0" cellpadding="0">
				  	<tr>
				  	  <td width="2"><img src="/images/filler.gif" width="2" height="38"></td>
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
		  <td width="2" bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
		  <td width="645" valign="top">
		  	<table border="0" cellspacing="0" cellpadding="0" width="645">
			  <tr>
			  	<td>
				  <table border="0" cellspacing="0" cellpadding="0">
				  	<tr>
				  	  <td width="2"><img src="/images/filler.gif" width="2" height="38"></td>
					  <td width="621" class="PageTitle" align="right">
					  	PRODUCT SALES SUMMARY
					  </td>
					  <td width="20" bgcolor="#A13846"><img src="/images/filler.gif" width="20" height="8"></td>
					  <td width="2"><img src="/images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td colspan="4"><img src="/images/filler.gif" width="1" height="20"></td>
					</tr>
				  </table>
				</td>
			  </tr>
<%
Call OpenDB()

'_____________________________________________________________________________________________
'CREATE THE PRODUCTS RECORDSET
SQL = "SELECT ProductID, Product FROM tblProduct WHERE Active = 1 AND CategoryID = 1"
	Set rsProducts = Conn.Execute(SQL)
	
	Do While Not rsProducts.EOF
		
		intProductID = rsProducts("ProductID")
		strProduct = rsProducts("Product")	
%>
              <tr>
			  	<td>
				  <table border="0" cellspacing="0" cellpadding="4" width="599" align="center">
<%
'_____________________________________________________________________________________________
'CREATE THE SALES RECORDSET
SQL = "SELECT O.OrderID, O.DateOrdered, C.Price, C.Quantity "
	SQL = SQL & "FROM (tblOrder O INNER JOIN relCartToOrder R ON O.OrderID = R.OrderID) "
	SQL = SQL & "INNER JOIN tblCart C ON R.CartID = C.CartID "
	SQL = SQL & "WHERE C.ProductID = " & intProductID & " ORDER BY O.DateOrdered DESC"
	Set rsCart = Conn.Execute(SQL)
	
If Not rsCart.EOF Then
%>
                    <tr>
                      <td colspan="3" style="padding-top:10px; font-size:16px; font-weight:bold;"><%=strProduct%></td>
                  	</tr>
                    <tr bgcolor="#CCCCCC">
					  <td width="349"><b>Date</b></td>
                      <td width="50"><b>Order #</b></td>
					  <td width="200"><b>Price</b></td>
					</tr>
<%
	myColor = "#E2E2E2"
	curTotalPrice = 0
	intTotalQuantity = 0
	
	Do While Not rsCart.EOF
	
		intOrderID = rsCart("OrderID")
		curPrice = rsCart("Price")
		dtDateOrdered = rsCart("DateOrdered")
		intQuantity = rsCart("Quantity")
		curPrice = curPrice * intQuantity
		curTotalPrice = curTotalPrice + curPrice
		intTotalQuantity = intTotalQuantity + intQuantity

		If myColor = "#E2E2E2" Then myColor = "#FFFFFF" Else myColor = "#E2E2E2"
%>
					<tr bgcolor="<%=myColor%>">
					  <td><%=dtDateOrdered%></td>
                      <td><a href="/orders/edit.asp?OrderID=<%=intOrderID%>"><%=intOrderID%></a></td>
					  <td><%=formatCurrency(curPrice,0)%></td>			  
					</tr>
<%
	rsCart.MoveNext
	Loop
%>
                    <tr style="background:none; margin-bottom:15px;">
					  <td colspan="3" style="text-align:right; font-size:14px; font-weight:bold;">Total: <%=formatCurrency(curTotalPrice,0)%><br />Quantity: <%=intTotalQuantity%></td>
					</tr>
<%
Else
%>
                    <tr>
                      <td style="padding:10px; font-size:16px; font-weight:bold; border:1px dashed #E2E2E2; text-align:right;"><%=strProduct%>: No Sales</td>
                  	</tr>
<%
End If
%>
				  </table>
				</td>
			  </tr>
<%
	rsProducts.MoveNext
	Loop
%>
			  <tr>
              	<td colspan="3">&nbsp;</td>
              </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
	<td bgcolor="#A13846" width="2"><img src="/images/filler.gif" width="2" height="1"></td>
  </tr>
  <tr>
  	<td colspan="3" bgcolor="#A13846" height="2"><img src="/images/filler.gif" width="2" height="2"></td>
  </tr>
</table>
<!--#include virtual="/incFooter.asp" -->

</body>
</html>
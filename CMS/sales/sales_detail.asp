<!--#include virtual="/includes/globalLib.asp"-->
<%
'_____________________________________________________________________________________________
'Get Variables
btnSubmit = Request("Submit")

dtStartDate = Request("StartDate")
dtEndDate = Request("EndDate")
	
If dtStartDate <> "" OR dtEndDate <> "" Then

	dtStartDate = Request("StartDate")
	dtEndDate = Request("EndDate")
	
Else
	
	dtStartDate = year(now) & "-01-01"
	
	tyear=year(date)
    tmonth=month(date)
    if tmonth <10 Then tmonth = "0" & tmonth
    tday=day(date)
    if tday<10 Then tday = "0" & tday
    tdate = tyear & "-" & tmonth & "-" & tday

	'dtEndDate = tdate
	dtEndDate = year(now) & "-12-31"
	
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
					  	SALES SUMMARY
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
			  <tr>
			  	<td>
				  <table border="0" cellspacing="0" cellpadding="0" style="margin-left:23px; margin-bottom:15px;">
                  	<form action="sales_detail.asp" method="post">
                    <tr>
                    	<td style="width:85px;"><input type="text" name="StartDate" value="<%=dtStartDate%>" style="width:75px;"></td>
                        <td style="width:85px;"><input type="text" name="EndDate" value="<%=dtEndDate%>" style="width:75px;"></td>
                        <td style="width:85px;"><input type="submit" name="mySubmit" value="Submit" style="width:100px;"></td>
                    </tr>
                    </form>
                  </table>
                </td>
              <tr>
			  	<td>
				  <table border="0" cellspacing="0" cellpadding="4" width="599" align="center">
<%
Call OpenDB()

'_____________________________________________________________________________________________
'CREATE THE PRODUCTS RECORDSET
SQL = "SELECT OrderID, PurchaseAmount, ShippingCost, DiscountAmount, TotalAmount, DateOrdered "
	SQL = SQL & " FROM  tblOrder "
	SQL = SQL & "WHERE DateOrdered >= '" & dtStartDate & "' AND DateOrdered <= '" & dtEndDate & " 11:59:59 PM' ORDER BY DateOrdered DESC"
	Set rsCart = Conn.Execute(SQL)
	
If Not rsCart.EOF Then
	
	myColor = "#E2E2E2"
	
	strHTML = "<tr bgcolor='#CCCCCC'>"_
		  & "<td><b>Date</b></td>"_
		  & "<td width='50'><b>Order #</b></td>"_
		  & "<td><b>Purchase Amount</b></td>"_
		  & "<td><b>Shipping</b></td>"_
		  & "<td><b>Discount</b></td>"_
		  & "<td><b>Total Amount</b></td>"_
		& "</tr>"
		
	Do While Not rsCart.EOF
	
		intOrderID = rsCart("OrderID")
		curPurchaseAmount = rsCart("PurchaseAmount")
		cShippingCost = rsCart("ShippingCost")
		curDiscountAmount = rsCart("DiscountAmount")
		curTotalAmount = rsCart("TotalAmount")
		dtDateOrdered = rsCart("DateOrdered")
		
		curPurchaseAmount_2 = curPurchaseAmount_2 + curPurchaseAmount
		cShippingCost_2 = cShippingCost_2 + cShippingCost
		curDiscountAmount_2 = curDiscountAmount_2 + curDiscountAmount
		curTotalAmount_2 = curTotalAmount_2 + curTotalAmount

		If myColor = "#E2E2E2" Then myColor = "#FFFFFF" Else myColor = "#E2E2E2"
		
			strHTML = strHTML & "<tr bgcolor='" & myColor & "'>"_
			  & "<td>" & dtDateOrdered & "</td>"_
			  & "<td><a href='/orders/edit.asp?OrderID=" & intOrderID & "'>" & intOrderID & "</a></td>"_
			  & "<td>" & formatCurrency(curPurchaseAmount,0) & "</td>"_
			  & "<td>" & formatCurrency(cShippingCost,0) & "</td>"_
			  & "<td>" & formatCurrency(curDiscountAmount,0) & "</td>"_
			  & "<td>" & formatCurrency(curTotalAmount,0) & "</td>"_
			& "</tr>"

	rsCart.MoveNext
	Loop

					strHTML2 = "<tr style='background-color:#CCCCCC; border-top:1px solid #000;'>"_
					  & "<td>&nbsp;</td>"_
                      & "<td>&nbsp;</td>"_
					  & "<td><b>Purchase Amount</b></td>"_
					  & "<td><b>Shipping</b></td>"_
					  & "<td><b>Discount</b></td>"_
					  & "<td><b>Total Amount</b></td>"_
					& "</tr>"_
                    & "<tr style='background-color:" & myColor & "; margin-bottom:15px;'>"_
					  & "<td>&nbsp;</td>"_
                      & "<td>&nbsp;</td>"_
					  & "<td style='font-size:14px; font-weight:bold;'>" & formatCurrency(curPurchaseAmount_2,0) & "</td>"_
					  & "<td style='font-size:14px; font-weight:bold;'>" & formatCurrency(cShippingCost_2,0) & "</td>"_
					  & "<td style='font-size:14px; font-weight:bold;'>" & formatCurrency(curDiscountAmount_2,0) & "</td>"_
					  & "<td style='font-size:14px; font-weight:bold;'>" & formatCurrency(curTotalAmount_2,0) & "</td>"_				  
					& "</tr>"
End If

Response.Write(strHTML2 & strHTML)
%>
				  </table>
				</td>
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
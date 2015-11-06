<!--#include virtual="/includes/globalLib.asp"-->
<%
Call OpenDB()
'_____________________________________________________________________________________________
'Get Variables
intOrderID = cInt(Request("OrderID"))
btnSubmit = Request("Submit")
intTrackingID = Request("TrackingID")

SQL = "SELECT B.FName, B.LName, B.Address, B.Address2, B.City, B.State, B.Zip, B.Email FROM " & _
	"tblBillingAddress B INNER JOIN tblOrder O ON B.BillingID = O.BillingID WHERE O.OrderID = " & intOrderID
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
	
SQL = "SELECT S.FName, S.LName, S.Address, S.Address2, S.City, S.State, S.Zip FROM " & _
	"tblShippingAddress S INNER JOIN tblOrder O ON S.ShippingID = O.ShippingID WHERE O.OrderID = " & intOrderID
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

CONST APIUser = "jdiehl@exacttarget.com"
CONST Password = "Scoobydoo1!"

CONST ListId = 2319
CONST EmailId = 2249

CONST strAPIUrl = "https://api.s4.exacttarget.com/integrate.aspx"

'_____________________________________________________________________________________________
'SEND TRACKING
If intOrderID > 0 AND btnSubmit <> "" Then
	
	strRequestXML = "<?xml version=" &chr(34) & "1.0" & chr(34)& "?><exacttarget><authorization>"_
		& "<username>"&APIUser&"</username>"_
		& "<password>"&Password&"</password>"_
	  & "</authorization>"_
	  & "<system>"_
		& "<system_name>triggeredsend</system_name>"_
		& "<action>add</action>"_
		& "<TriggeredSend xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns='http://exacttarget.com/wsdl/partnerAPI'>"_
		  & "<TriggeredSendDefinition>"_
			& "<CustomerKey>Shipping</CustomerKey>"_
		  & "</TriggeredSendDefinition>"_
		  & "<Subscribers>"_
			& "<SubscriberKey>"&strB_Email&"</SubscriberKey>"_
			& "<EmailAddress>"&strB_Email&"</EmailAddress>"_
			& "<Attributes>"_
			  & "<Name>First Name</Name>"_
			  & "<Value>"&strB_FName&"</Value>"_
			& "</Attributes>"_
			& "<Attributes>"_
			  & "<Name>Last Name</Name>"_
			  & "<Value>"&strB_LName&"</Value>"_
			& "</Attributes>"_
			& "<Attributes>"_
			  & "<Name>trackingID</Name>"_
			  & "<Value>"&intTrackingID&"</Value>"_
			& "</Attributes>"_
			& "<Attributes>"_
			  & "<Name>Id</Name>"_
			  & "<Value></Value>"_
			& "</Attributes>"_
		  & "</Subscribers>"_
		& "</TriggeredSend>"_
	  & "</system>"_
	& "</exacttarget>"

	Set objSXH = server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objSXH.open "POST", strAPIUrl, false

		' Post the XML body
		objSXH.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		objSXH.send ("qf=xml&xml=" & Server.URLEncode(strRequestXML))
		
		'Check the status of the request and response accordingly.
		If objSXH.status = 200 Then

			Set objResponseXML = objSXH.responseXML

			'look for subscriber_description
			Set objNodeList = objResponseXML.GetElementsByTagName("triggered_send_description")
				intNodeCount = objNodeList.length

				If intNodeCount = 0 Then
					'handle the error.
					errMessage = "There was an error saving your preferences. Please try again later.<br />Nodecount: " & NodeCount
				Else
					errMessage = "Success: " & objNodeList.Item(0).Text
				End If
				
			Set objNodeList = Nothing

			Set objResponseXML = Nothing

		Else
			' handle the error information
			errMessage = "There was an error saving your preferences. Please try again later.<br />" & objXMLHttp.StatusText
		End If

		Set objSXH = nothing
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
					  	ORDER :: INFO
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
<form action="edit.asp?Submit=True&OrderID=<%=intOrderID%>" method="post">
              <tr>
              	<td>
                	<table border="0" cellspacing="0" cellpadding="4" width="599" align="center">
                    	<tr><td>Confirm Shipping <input type="text" name="trackingID" width="350"> <input type="submit" value="submit"></td></tr>
                        <tr><td><%=errMessage%></td></tr>
                    </table>
                </td>
              </tr>
</form>
			  <tr>
			  	<td>
				  <table border="0" cellspacing="0" cellpadding="4" width="599" align="center">
					<tr bgcolor="#CCCCCC">
					  <td width="50"><b>Order #</b></td>
					  <td width="341"><b>Product</b></td>
					  <td width="70"><b>Size</b></td>
					  <td width="40"><b>Price</b></td>
					  <td width="60"><b>Shipping</b></td>
					  <td width="30"><b>Qty</b></td>
					</tr>
<%
Call OpenDB()

'_____________________________________________________________________________________________
'CREATE THE PRODUCTS RECORDSET
SQL = "SELECT C.CartID, P.Product, C.Price, S.Size, Y.Style, C.Quantity, R.PurchaseAmount, R.ShippingCost, R.DiscountAmount, R.TotalAmount "
	SQL = SQL & " FROM (((((tblCart C INNER JOIN tblProduct P ON C.ProductID = P.ProductID) "
	SQL = SQL & "INNER JOIN tblProductSize S ON C.ProductSizeID = S.ProductSizeID) "
	SQL = SQL & "INNER JOIN tblProductStyle Y ON C.ProductStyleID = Y.ProductStyleID)"
	SQL = SQL & "INNER JOIN relCartToOrder O ON C.CartID = O.CartID)"
	SQL = SQL & "INNER JOIN tblOrder R ON O.OrderID = R.OrderID) "
	SQL = SQL & "WHERE O.OrderID = " & intOrderID
	Set rsCart = Conn.Execute(SQL)
	
If Not rsCart.EOF Then
	
	myColor = "#E2E2E2"
	
	curPurchaseAmount = rsCart("PurchaseAmount")
	cShippingCost = rsCart("ShippingCost")
	curDiscountAmount = rsCart("DiscountAmount")
	curTotalAmount = rsCart("TotalAmount")
	
	Do While Not rsCart.EOF
	
		intCartID = rsCart("CartID")
		strProduct = rsCart("Product")
		strStyle = rsCart("Style")
		strSize = rsCart("Size")
		curPrice = rsCart("Price")
		intQuantity = rsCart("Quantity")

		If myColor = "#E2E2E2" Then myColor = "#FFFFFF" Else myColor = "#E2E2E2"
%>
					<tr bgcolor="<%=myColor%>">
					  <td><%=intOrderID%></td>
					  <td><%=strProduct%></td>
					  <td><%=strSize%></td>
					  <td><%=formatCurrency(curPrice,0)%></td>
					  <td>&nbsp;</td>
					  <td align="center"><%=intQuantity%></td>				  
					</tr>
<%
	rsCart.MoveNext
	Loop
	
End If
%>
					<tr style="padding:5px;">
					  <td colspan="6" valign="top" align="right">
					  	<table border="0" cellspacing="0" cellpadding="0">
						  <tr>
						  	<td style="text-align:right;">Sub Total: </td>
							<td  style="text-align:right; width:100px;"><%=formatCurrency(curPurchaseAmount,2)%></td>
						  </tr>
						  <tr>
						  	<td style="text-align:right;">Discount: </td>
							<td style="text-align:right;"><%=formatCurrency(curDiscountAmount,2)%></td>
						  </tr>
						  <tr>
						  	<td style="padding-bottom:12px; text-align:right;">Shipping: </td>
						  	<td style="padding-bottom:12px; text-align:right;"><%=formatCurrency(cShippingCost,2)%></td>
						  </tr>
						  <tr style="font-weight:bold;">
						  	<td style="text-align:right;">Total amount: </td>
							<td style="text-align:right;"><%=formatCurrency(curTotalAmount)%></td>
						  </tr>
						</table>
					  </td>
					</tr>
				  </table>
				  <table border="0" cellspacing="0" cellpadding="0" width="599" align="center">
                  <tr>
                    <td><strong>email:</strong> <input type="text" value="<%=strB_Email%>" style="width:350px;"></td>
                  </tr>
				  <tr>
                    <td height="10"><img src="/images/filler.gif" width="1" height="10"></td>
                  </tr>
				  <tr>
                    <td height="1" bgcolor="#C84848"><img src="/images/filler.gif" width="1" height="1"></td>
                  </tr>
				  <tr>
                    <td><strong>billing address</strong></td>
                  </tr>
                  <tr>
                    <td height="8"><img src="/images/filler.gif" width="1" height="8"></td>
                  </tr>
                  <tr>
                    <td><%=strB_FName%>&nbsp;<%=strB_LName%><br><%=strB_Address%><%If strB_Address2 <> "" Then Response.Write("<br>" & strB_Address2)%><br><%=strB_City%>, <%=strB_State%>&nbsp;&nbsp;<%=strB_Zip%></td>
                  </tr>
				  <tr>
                    <td height="10"><img src="/images/filler.gif" width="1" height="10"></td>
                  </tr>
				  <tr>
                    <td height="1" bgcolor="#C84848"><img src="/images/filler.gif" width="1" height="1"></td>
                  </tr>
				  <tr>
                    <td><strong>shipping address</strong></td>
                  </tr>
                  <tr>
                    <td height="8"><img src="/images/filler.gif" width="1" height="8"></td>
                  </tr>
                  <tr>
                    <td><%=strS_FName%>&nbsp;<%=strS_LName%><br><%=strS_Address%><%If strS_Address2 <> "" Then Response.Write("<br>" & strS_Address2)%><br><%=strS_City%>, <%=strS_State%>&nbsp;&nbsp;<%=strS_Zip%></td>
                  </tr>
				  <tr>
                    <td height="10"><img src="/images/filler.gif" width="1" height="10"></td>
                  </tr>
				  <tr>
                    <td height="1" bgcolor="#C84848"><img src="/images/filler.gif" width="1" height="1"></td>
                  </tr>
				  <tr>
                    <td><strong>credit card</strong></td>
                  </tr>
                  <tr>
                    <td height="8"><img src="/images/filler.gif" width="1" height="8"></td>
                  </tr>
<%
SQL = "SELECT P.CardType, P.CardNumber, P.ExpMo, P.ExpYear FROM " & _
	"tblPayment P INNER JOIN tblOrder O ON P.PaymentID = O.PaymentID WHERE O.OrderID = " & intOrderID
	Set rsPayment = Conn.Execute(SQL)
	
	strCardType = rsPayment("CardType")
	strCardNumber = rsPayment("CardNumber")
	strCardNumber = right(strCardNumber,4)
	strCardNumber = "XXXX-XXXX-XXXX-" & strCardNumber
	strExpMo = rsPayment("ExpMo")
	strExpYear = rsPayment("ExpYear")
	
	rsPayment.Close
	Set rsPayment = Nothing

Call CloseDB()
%>
                  <tr>
                    <td><%=strB_FName%>&nbsp;<%=strB_LName%><br>
                      <%=strCardType%><br>
                      <%=strCardNumber%><br>
                      <%=strExpMo%>/<%=strExpYear%></td>
                  </tr>
				  <tr>
                    <td height="10"><img src="/images/filler.gif" width="1" height="10"></td>
                  </tr>
				  <tr>
				  	<td align="center" style="padding-bottom:15px;"><a href="print.asp?OrderID=<%=intOrderID%>">- PRINT RECEIPT -</a></td>
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
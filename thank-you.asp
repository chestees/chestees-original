<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="/includes/globalLib.asp" -->
<%
intOrderID = Request.Cookies(cSiteName)("OrderID")
If intOrderID = 0 Then
	response.redirect("/")
End If
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Confirm Chestees Order</title>
<script src="/js/jquery-1.4.2.min.js" type="text/javascript"></script>
<script src="/js/functions.js" type="text/javascript"></script>
<link rel="stylesheet" href="/css/style.css">
<link rel="shortcut icon" href="/images/favicon.ico" type="image/x-icon">
</head>
<body>
<script type="text/javascript">
  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-3725315-1']);
  _gaq.push(['_setDomainName', '.chestees.com']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();
</script>
<% Call Header() %>
<div class="Main">
	<div class="TwoColumn_Left Module">
    	<H1>Thank You! Your order has been processed.</H1>

        <div style="margin-bottom:10px;">
        	<div style="float:left; width:339px;">
            	Your order number is: <b><%=intOrderID%></b><br /><br />
            	What happens next:<br />
				- We'll email you an order confirmation within a few minutes.<br />
                - We'll email you again once we've shipped your order.<br /><br />
        	<a class="medium red button" href="javascript:print(document);">print this page for your record</a></div>
        	<div style="float:right; width:212px; margin-left:10px; text-align:right;">
            	<a class="btnFacebook" href="http://www.facebook.com/chestees" target="_blank"><span>"LIKE" us on Facebook</span></a>
            	<a class="btnTwitter" href="http://twitter.com/chestees" target="_blank"><span>Follow us on Twitter why don't cha...</span></a>
            	<div class="clear"></div>
            </div>
        	<div class="clear"></div>
        </div>
        <!-- CART INFO -->
        <div class="Cart_Header">
            <div class="Cart_Col1">T-Shirt</div>
            <div class="Cart_Col2-3">Size</div>
            <div class="Cart_Col2-3">Price</div>
            <div class="Cart_Col4">Qty</div>
            <div class="clear"></div>
        </div>
<%
Call OpenDB()
'_____________________________________________________________________________________________
'CREATE THE PRODUCTS RECORDSET
SQL = "SELECT C.CartID, P.Product, C.Price, S.SizeAbbr, Y.Style, C.Quantity, R.PurchaseAmount, R.ShippingCost, R.DiscountAmount, R.TotalAmount "
	SQL = SQL & "FROM (((((tblCart C INNER JOIN tblProduct P ON C.ProductID = P.ProductID) "
	SQL = SQL & "INNER JOIN tblProductSize S ON C.ProductSizeID = S.ProductSizeID) "
	SQL = SQL & "INNER JOIN tblProductStyle Y ON C.ProductStyleID = Y.ProductStyleID)"
	SQL = SQL & "INNER JOIN relCartToOrder O ON C.CartID = O.CartID)"
	SQL = SQL & "INNER JOIN tblOrder R ON O.OrderID = R.OrderID) "
	SQL = SQL & "WHERE O.OrderID = " & intOrderID
	Set rsCart = Conn.Execute(SQL)
	
If Not rsCart.EOF Then

	myColor = "#E2E2E2"
	
	curPurchaseAmount = rsCart("PurchaseAmount")
	curShippingCost = rsCart("ShippingCost")
	intDiscountAmount = rsCart("DiscountAmount")
	curTotalAmount = rsCart("TotalAmount")
	
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
        <div class="Cart_Header_Mod" style="float:background-color:<%=myColor%>;">
            <div class="Cart_Col1"><b><%=strProduct%></b><br /><span class="SmallText">Color: <%=strStyle%></span></div>
            <div class="Cart_Col2-3"><%=strSizeAbbr%></div>
            <div class="Cart_Col2-3"><%=formatCurrency(curPrice,0)%></div>
            <div class="Cart_Col4"><%=intQuantity%></div>
            <div class="clear"></div>
        </div>
<%
	rsCart.MoveNext
	Loop
	
End If
%>
        <!-- END CART INFO -->
        
        <!-- BILLING -->
<%	
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
%>
		<div class="Checkout_Header">Billing Address</div>
        <div style="margin:10px auto auto 25px;">
            <%=strB_FName%>&nbsp;<%=strB_LName%><br><%=strB_Address%><%If strB_Address2 <> "" Then Response.Write("<br>" & strB_Address2)%><br><%=strB_City%>, <%=strB_State%>&nbsp;&nbsp;<%=strB_Zip%>
        </div>
        <!-- END BILLING -->
    
        <!-- SHIPPING -->
<%	
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
        <div class="Checkout_Header">Shipping Address</div>                       
        <div style="margin:10px auto auto 25px;">
            <%=strS_FName%>&nbsp;<%=strS_LName%><br><%=strS_Address%><%If strS_Address2 <> "" Then Response.Write("<br>" & strS_Address2)%><br><%=strS_City%>, <%=strS_State%>&nbsp;&nbsp;<%=strS_Zip%>
        </div>  
        <!-- END SHIPPING -->
        
        <!-- PAYMENT -->
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
        <div class="Checkout_Header">Credit Card Information</div>
        <div style="margin:10px auto auto 25px;">
            <%=strB_FName%>&nbsp;<%=strB_LName%><br>
            <%=strCardType%><br>
            <%=strCardNumber%><br>
            <%=strExpMo%>/<%=strExpYear%><br>
            <%=strCVV%>
        </div>
        <!-- END PAYMENT -->
        
        <!-- TOTALS -->
<%
If strCouponCode <> "" Then

	SQL = "SELECT Discount, DollarAmount FROM tblCouponCode WHERE CouponCode = " & SQLEncode(strCouponCode)
		Set rsCoupon = Conn.Execute(SQL)
	
	If Not rsCoupon.EOF Then
		intDiscount = rsCoupon("Discount")
		intDollarAmount = rsCoupon("DollarAmount")
		If intDiscount > 0 Then
			curDiscountAmount = curTotal * intDiscount/100
			strDiscount = intDiscount & "% OFF"
		ElseIf intDollarAmount > 0 Then
			curDiscountAmount = intDollarAmount
			strDiscount = "$" & curDiscountAmount & " OFF"
		End If
		
		curSubTotal = curTotal - curDiscountAmount
	Else
		curSubTotal = curTotal
		strDiscount = "Invalid Code"
	End If
	rsCoupon.Close
	Set rsCoupon = Nothing
End If

Call CloseDB()
%>
        <div class="PaymentSummary Module">
            <div style="float:left; width:494px; margin-top:5px; margin-right:5px; text-align:right;">Subtotal:</div>
            <div style="float:right; width:80px; margin-top:5px; text-align:right;"><%=formatCurrency(curPurchaseAmount,2)%></div>
            <div class="clear"></div>

            <%If strCouponCode <> "" AND strDiscount <> "Invalid Code" Then%>
            <div style="float:left; width:494px; margin-top:5px; margin-right:5px; text-align:right; color:#b93636; vertical-align:top;">Discount:<br /><span style="font-size:11px;">(<%=strCouponCode%> = <%=strDiscount%>)</span></div>
            <div style="float:right; width:80px; margin-top:5px; text-align:right; color:#b93636; vertical-align:top;"><%=formatCurrency(curDiscountAmount,2)%></div>
            <div class="clear"></div>
            <%End If%>

            <div style="float:left; width:494px; margin-top:5px; margin-right:5px; text-align:right;">Shipping:</div>
            <div style="float:right; width:80px; margin-top:5px; text-align:right;"><%=formatCurrency(curShippingCost,2)%></div>
            <div class="clear"></div>

            <div style="float:left; width:494px; margin-top:5px; margin-right:5px; text-align:right; font-weight:bold;">Total:</div>
            <div style="float:right; width:80px; margin-top:5px; text-align:right; font-weight:bold;"><%=formatCurrency(curTotalAmount)%></div>
            <div class="clear"></div>
        </div>    
        <!-- END TOTALS -->
    </div>
	<div class="Seals Module">
        <div class="Seals_Row_Center" style="padding-top:0;">
            <span id="siteseal"><script type="text/javascript" src="https://seal.godaddy.com/getSeal?sealID=lBM8wbUtZODPoQAPq2qArFpShU8Xya79V4QUwoJnAgxQWI4T3PGuEk"></script></span>
        </div>
        <div class="Seals_Row_Center">
            <a href="https://www.paypal.com/us/verified/pal=jdiehl%40jasondiehl%2ecom" target="_blank"><img src="https://www.paypal.com/en_US/i/icon/verification_seal.gif" border="0" alt="Official PayPal Seal"></a>
        </div>
        <div class="Seals_Row_Left">Orders arrive within 5-7 business days of the order being placed.</div>
        <div class="Seals_Row_Left">All orders are shipped First Class with the USPS</div>
        <div class="Seals_Row_Left">Returns are accepted given the product is returned unwashed and unworn.</div>
    </div>
    <div class="clear"></div>
</div>
<!--#include virtual="/incFooter_s.asp" -->
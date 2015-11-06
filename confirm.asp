<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="/includes/globalLib.asp" -->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Confirm Chestees Order</title>
<script src="/js/jquery-1.4.2.min.js" type="text/javascript"></script>
<script src="/js/jquery-ui-1.8.2.custom.min.js" type="text/javascript"></script>
<script src="/js/confirm.js" type="text/javascript"></script>
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
<div id="MessageBar"></div>
<% Call Header() %>
<div class="Main">
	<div class="TwoColumn_Left Module">
    
        <H1>Confirm Order</H1>          
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
'CREATE THE PRODUCTS SIZES RECORDSET
SQL = "SELECT C.CartID, P.Product, C.Price, S.SizeAbbr, Y.Style, C.Quantity FROM (((tblCart C INNER JOIN tblProduct P ON C.ProductID = P.ProductID) "
	SQL = SQL & "INNER JOIN tblProductSize S ON C.ProductSizeID = S.ProductSizeID) "
	SQL = SQL & "INNER JOIN tblProductStyle Y ON C.ProductStyleID = Y.ProductStyleID) "
	SQL = SQL & "WHERE C." & varBuyer() & "ID = " & cBuyerID()
	SQL = SQL & " AND C.Purchased = 0"
	Set rsCart = Conn.Execute(SQL)
	
If Not rsCart.EOF Then
	
	curTotal = 0
	curTotalShipping = 0
	intTotalQuantity = 0
	myColor = "#E2E2E2"
	
	Do While Not rsCart.EOF
	
		intCartID = rsCart("CartID")
		strProduct = rsCart("Product")
		strStyle = rsCart("Style")
		strSizeAbbr = rsCart("SizeAbbr")
		curPrice = rsCart("Price")
		intQuantity = rsCart("Quantity")
		intTotalQuantity = intTotalQuantity + intQuantity
		curTotal = curTotal+curPrice*intQuantity
		
		If intTotalQuantity >= cFreeShippingNum Then
			curShippingCost = 0
		Else 
			curShippingCost = cShippingCost
		End If
		If myColor = "#E2E2E2" Then myColor = "#FFFFFF" Else myColor = "#E2E2E2"
%>
        <div class="Cart_Header_Mod" style="background-color:<%=myColor%>;">
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
SQL = "SELECT FName, LName, Address, Address2, City, State, Zip, Email FROM tblBillingAddress WHERE Lock = 0 AND " & varBuyer() & "ID = " & cBuyerID()
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
        <div style="float:left; margin:10px auto auto 25px; width:399px;">
            <%=strB_FName%>&nbsp;<%=strB_LName%><br /><%=strB_Address%><%If strB_Address2 <> "" Then Response.Write("<br />" & strB_Address2)%><br /><%=strB_City%>, <%=strB_State%>&nbsp;&nbsp;<%=strB_Zip%>
        </div>
        <div style="float:right; margin-top:10px; width:150px;"><a class="small red button" href="/checkout/">edit this information</a></div>
        <div class="clear"></div>
        <!-- END BILLING -->

        <!-- SHIPPING -->
<%	
SQL = "SELECT FName, LName, Address, Address2, City, State, Zip FROM tblShippingAddress WHERE Lock = 0 AND " & varBuyer() & "ID = " & cBuyerID()
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
        <div style="float:left; margin:10px auto auto 25px; width:399px;">
            <%=strS_FName%>&nbsp;<%=strS_LName%><br /><%=strS_Address%><%If strS_Address2 <> "" Then Response.Write("<br />" & strS_Address2)%><br /><%=strS_City%>, <%=strS_State%>&nbsp;&nbsp;<%=strS_Zip%>
        </div>
        <div style="float:right; margin-top:10px; width:150px;"><a class="small red button" href="/checkout/">edit this information</a></div>
        <div class="clear"></div>  
        <!-- END SHIPPING -->
        
        <!-- PAYMENT -->
<%
SQL = "SELECT CardType, CardNumber, ExpMo, ExpYear, CouponCode, CVV FROM tblPayment WHERE Lock = 0 AND " & varBuyer() & "ID = " & cBuyerID()
	Set rsPayment = Conn.Execute(SQL)
	
	strCardType = rsPayment("CardType")
	strCardNumber = rsPayment("CardNumber")
	strCardNumber = right(strCardNumber,4)
	strCardNumber = "XXXX-XXXX-XXXX-" & strCardNumber
	strExpMo = rsPayment("ExpMo")
	strExpYear = rsPayment("ExpYear")
	strCVV = rsPayment("CVV")
	strCouponCode = rsPayment("CouponCode")
	
	rsPayment.Close
	Set rsPayment = Nothing
%>
        <div class="Checkout_Header">Credit Card Information</div>
        <div style="float:left; margin:10px auto auto 25px; width:399px;">
            <%=strB_FName%>&nbsp;<%=strB_LName%><br />
            <%=strCardType%><br />
            <%=strCardNumber%><br />
            <%=strExpMo%>/<%=strExpYear%><br />
            <%=strCVV%><br /><br />
            <%If strCouponCode <> "" Then response.Write("Coupon Code: " & strCouponCode)%>
        </div>    
        <div style="float:right; margin-top:10px; width:150px;"><a class="small red button" href="/checkout/">edit this information</a></div>
        <div class="clear"></div>
        <!-- END PAYMENT -->
        
        <!-- TOTALS -->
<%If strCouponCode <> "" Then

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
Else
	curSubTotal = curTotal
End If

Call CloseDB()
%>
        <div class="PaymentSummary Module">
            <div style="float:left; width:494px; margin-top:5px; margin-right:5px; text-align:right;">Subtotal:</div>
            <div style="float:right; width:80px; margin-top:5px; text-align:right;"><%=formatCurrency(curTotal,2)%></div>
            <div class="clear"></div>

            <%If strCouponCode <> "" Then%>
            <div style="float:left; width:494px; margin-top:5px; margin-right:5px; text-align:right; color:#b93636; vertical-align:top;">Discount:<br /><span style="font-size:11px;">(<%=strCouponCode%> = <%=strDiscount%>)</span></div>
            <div style="float:right; width:80px; margin-top:5px; text-align:right; color:#b93636; vertical-align:top;"><%=formatCurrency(curDiscountAmount,2)%></div>
            <div class="clear"></div>
            <%End If%>

            <div style="float:left; width:494px; margin-top:5px; margin-right:5px; text-align:right;">Shipping:</div>
            <div style="float:right; width:80px; margin-top:5px; text-align:right;"><%=formatCurrency(curShippingCost,2)%></div>
            <div class="clear"></div>

            <div style="float:left; width:494px; margin-top:5px; margin-right:5px; text-align:right; font-weight:bold;">Total:</div>
            <div style="float:right; width:80px; margin-top:5px; text-align:right; font-weight:bold;"><%=formatCurrency(curSubTotal+curShippingCost)%></div>
            <div class="clear"></div>
        </div>    
        <!-- END TOTALS -->

        <input type="hidden" id="PurchaseAmount" value="<%=curTotal%>">
        <input type="hidden" id="ShippingCost" value="<%=curShippingCost%>">
        <input type="hidden" id="CouponCode" value="<%=strCouponCode%>">
        <input type="hidden" id="TotalAmount" value="<%=curSubTotal+curShippingCost%>">
        
        <div class="Submit_Bar">
            <input type="button" id="mySubmit" class="large red button" name="Submit" value="Place My Order">
        </div>
	</div>
	<div class="Seals">
        <div class="Seals_Row_Center" style="padding-top:0;">
            <span id="siteseal"><script type="text/javascript" src="https://seal.godaddy.com/getSeal?sealID=lBM8wbUtZODPoQAPq2qArFpShU8Xya79V4QUwoJnAgxQWI4T3PGuEk"></script></span>
        </div>
        <div class="Seals_Row_Left">Orders arrive within 5-7 business days of the order being placed.</div>
        <div class="Seals_Row_Left">All orders are shipped First Class with the USPS</div>
        <div class="Seals_Row_Left">Returns are accepted given the product is returned unwashed and unworn.</div>
    </div>
    <div class="clear"></div>
</div>
<!--#include virtual="/incFooter_s.asp" -->
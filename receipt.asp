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
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Chestees Order</title>
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
<!-- START MAIN AREA CONTENT -->
<div style="margin:0 auto; width:429px; text-align:left; padding-bottom:25px;">
    <div style="text-align:center; width:auto; margin-bottom:10px;"><img src="/images/chestees_Logo.png" alt="Chestees Funny T-Shirts" title="Chestees Funny T-Shirts"></div>
    <div style="float:left; border:1px solid #000000; background-color:#FFFFFF; padding:15px;">
        <div style="margin:0 0 10px 0; text-align:center;">[ <a style="color:#b93636;" href="javascript:print(document);">print this receipt</a> ]</div>
        <H1>Order Confirmation</H1>

        <div id="myForm">

            <!-- CART INFO -->
            <div style="float:left; width:397px; background-color:#CCCCCC; border-top:1px solid #b93636; border-bottom:1px solid #C84848;">
                <div style="width:232px; float:left; font-weight:bold; padding:5px 0 5px 0;">T-Shirt</div>
                <div style="width:60px; float:left; font-weight:bold; text-align:center; padding:5px 0 5px 0;">Size</div>
                <div style="width:60px; float:left; font-weight:bold; text-align:center; padding:5px 0 5px 0;">Price</div>
                <div style="width:35px; float:left; font-weight:bold; text-align:center; padding:5px 0 5px 0;">Qty</div>
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
            <div style="float:left; width:397px; background-color:<%=myColor%>;">
                <div style="width:232px; float:left; vertical-align:top; padding:5px 0 5px 0;"><b><%=strProduct%></b><br><span style="font-size:12px;"><%=strStyle%></span></div>
                <div style="width:60px; float:left; vertical-align:top; text-align:center; padding:5px 0 5px 0;"><%=strSizeAbbr%></div>
                <div style="width:60px; float:left; vertical-align:top; text-align:center; padding:5px 0 5px 0;"><%=formatCurrency(curPrice,0)%></div>
                <div style="width:35px; float:left; vertical-align:top; text-align:center; padding:5px 0 5px 0;"><%=intQuantity%></div>
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
            <div style="float:left; padding:5px; border:1px solid #000; width:385px; margin-bottom:10px; margin-top:10px;">
                <div style="float:left; width:375px; padding:5px; background-color:#CCCCCC; border-top:1px solid #b93636; border-bottom:1px solid #C84848; font-weight:bold;">
                    Billing Address
                </div>

                <div style="float:left; width:375px;">
                    <div style="float:left; width:200px; margin-top:10px;">
                        <%=strB_FName%>&nbsp;<%=strB_LName%><br><%=strB_Address%><%If strB_Address2 <> "" Then Response.Write("<br>" & strB_Address2)%><br><%=strB_City%>, <%=strB_State%>&nbsp;&nbsp;<%=strB_Zip%>
                    </div>
                </div>
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
            <div style="float:left; padding:5px; border:1px solid #000; width:385px; margin-bottom:10px;">                    
                <div style="float:left; width:375px; padding:5px; background-color:#CCCCCC; border-top:1px solid #b93636; border-bottom:1px solid #C84848; font-weight:bold;">
                    Shipping Address
                </div>                       

                <div style="float:left; width:375px;">
                    <div style="float:left; width:200px; margin-top:10px;">
                        <%=strS_FName%>&nbsp;<%=strS_LName%><br><%=strS_Address%><%If strS_Address2 <> "" Then Response.Write("<br>" & strS_Address2)%><br><%=strS_City%>, <%=strS_State%>&nbsp;&nbsp;<%=strS_Zip%>
                    </div>
                </div>    
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
            <div style="float:left; padding:5px; border:1px solid #000; width:385px; margin-bottom:10px;">
            
                <div style="float:left; width:375px; padding:5px; background-color:#CCCCCC; border-top:1px solid #b93636; border-bottom:1px solid #C84848; font-weight:bold;">
                    Credit Card Information
                </div>

                <div style="float:left; width:375px;">
                    <div style="float:left; width:200px; margin-top:10px;">
                        <%=strB_FName%>&nbsp;<%=strB_LName%><br>
                        <%=strCardType%><br>
                        <%=strCardNumber%><br>
                        <%=strExpMo%>/<%=strExpYear%><br>
                        <%=strCVV%>
                    </div>
                </div>    
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
			curDiscountAmount = curPurchaseAmount * intDiscount/100
			strDiscount = intDiscount & "% OFF"
		ElseIf intDollarAmount > 0 Then
			curDiscountAmount = intDollarAmount
			strDiscount = "$" & curDiscountAmount & " OFF"
		End If
		
		'curSubTotal = curTotal - curDiscountAmount
	Else
		'curSubTotal = curTotal
	End If
	rsCoupon.Close
	Set rsCoupon = Nothing
'Else
	'curSubTotal = curTotal
End If

Call CloseDB()
%>
            <div style="float:right;">
                <div style="float:left; width:145px; margin-top:5px; margin-right:5px; text-align:right;">Subtotal:</div>
                <div style="float:right; width:80px; margin-top:5px; text-align:right;"><%=formatCurrency(curPurchaseAmount,2)%></div>
                <div class="clear"></div>
            </div>
            <div style="float:right;">
                <%If strCouponCode <> "" Then%>
                <div style="float:left; width:300px; margin-top:5px; margin-right:5px; text-align:right; color:#b93636; vertical-align:top;">Discount:<br /><span style="font-size:11px;">(<%=strCouponCode%> = <%=strDiscount%>)</span></div>
                <div style="float:right; width:80px; margin-top:5px; text-align:right; color:#b93636; vertical-align:top;"><%=formatCurrency(curDiscountAmount,2)%></div>
                <%End If%>
                <div class="clear"></div>
            </div>
            <div style="float:right;">
                <div style="float:left; width:145px; margin-top:5px; margin-right:5px; text-align:right;">Shipping:</div>
                <div style="float:right; width:80px; margin-top:5px; text-align:right;"><%=formatCurrency(curShippingCost,2)%></div>
                <div class="clear"></div>
            </div>
            <div style="float:right;">
                <div style="float:left; width:145px; margin-top:5px; margin-right:5px; text-align:right; font-weight:bold;">Total:</div>
                <div style="float:right; width:80px; margin-top:5px; text-align:right; font-weight:bold;"><%=formatCurrency(curTotalAmount)%></div>
                <div class="clear"></div>
            </div>
            <!-- END TOTALS -->
        </div>
    </div>
</div>

</body>
</html>
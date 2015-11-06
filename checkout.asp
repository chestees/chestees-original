<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="/includes/globalLib.asp" -->
<!--#include virtual="/includes/adovbs.inc" -->
<%
OpenDB()
SQL = "SELECT FName, LName, Address, Address2, City, State, Zip, Email FROM tblBillingAddress WHERE Lock = 0 AND " & varBuyer() & "ID = " & cBuyerID()
	Set rsBilling = Conn.Execute(SQL)
	
	If Not rsBilling.EOF Then
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
	End If
	
SQL = "SELECT FName, LName, Address, Address2, City, State, Zip FROM tblShippingAddress WHERE Lock = 0 AND " & varBuyer() & "ID = " & cBuyerID()
	Set rsShipping = Conn.Execute(SQL)
	
	If Not rsShipping.EOF Then
		strS_FName = rsShipping("FName")
		strS_LName = rsShipping("LName")
		strS_Address = rsShipping("Address")
		strS_Address2 = rsShipping("Address2")
		strS_City = rsShipping("City")
		strS_State = rsShipping("State")
		strS_Zip = rsShipping("Zip")
		
		rsShipping.Close
		Set rsShipping = Nothing
	End If
	
SQL = "SELECT CardType, CardNumber, ExpMo, ExpYear, CVV, CouponCode FROM tblPayment WHERE Lock = 0 AND " & varBuyer() & "ID = " & cBuyerID()
	Set rsPayment = Conn.Execute(SQL)
	
	If Not rsPayment.EOF Then
		strCardType = rsPayment("CardType")
		strCardNumber = rsPayment("CardNumber")
		strExpMo = rsPayment("ExpMo")
		strExpYear = rsPayment("ExpYear")
		strCVV = rsPayment("CVV")
		strCouponCode = rsPayment("CouponCode")
		
		rsPayment.Close
		Set rsPayment = Nothing
	End If
CloseDB()
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Secure Checkout</title>
<link rel="stylesheet" href="/css/style.css">
<link rel="shortcut icon" href="/images/favicon.ico" type="image/x-icon">
<script src="/js/jquery-1.4.2.min.js" type="text/javascript"></script>
<script src="/js/jquery-ui-1.8.2.custom.min.js" type="text/javascript"></script>
<script src="/js/checkout.js" type="text/javascript"></script>
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
	<div class="TwoColumn_Left">
        <H1>Secure Checkout</H1>
        
        <div id="ErrorRow"></div>
<form>
        <!-- BILLING -->
        <div class="Checkout_Header">Billing Address</div>
        <div class="Checkout_Field">
            <div style="float:left;">
                First Name <span class="Required">*required</span><br><input name="B_FName" id="B_FName" type="text" class="rounded-glow w270" value="<%=strB_FName%>">
            </div>
            <div style="float:right;">
                Last Name <span class="Required">*required</span><br>
                  <input name="B_LName" id="B_LName" type="text" class="rounded-glow w270" value="<%=strB_LName%>">
            </div>
            <div class="clear"></div>
        </div>
        <div class="Checkout_Field">       
            <div style="float:left;">
                Email <span class="Required">*required</span><br><input name="B_Email" id="B_Email" type="text" class="rounded-glow w270" value="<%=strB_Email%>">
            </div>
            <div style="float:right;">
                Confirm Email <span class="Required">*required</span><br><input name="B_Email_Confirm" id="B_Email_Confirm" type="text" class="rounded-glow w270" value="<%=strB_Email%>">
            </div>
            <div class="clear"></div>
        </div>
        <div class="Checkout_Field">
            Address <span class="Required">*required</span><br><input name="B_Address" id="B_Address" type="text" class="rounded-glow w350" value="<%=strB_Address%>">
        </div>
        <div class="Checkout_Field">
            Address 2<br><input name="B_Address2" id="B_Address2" type="text" class="rounded-glow w350" value="<%=strB_Address2%>">
        </div>
        <div class="Checkout_Field">
            <div style="float:left;">
                City <span class="Required">*required</span><br><input name="B_City" id="B_City" type="text" class="rounded-glow w150" value="<%=strB_City%>">
            </div>
            <div style="float:left; margin-left:25px;">
                State <span class="Required">*required</span><br>
                <select name="B_State" id="B_State" class="rounded-glow" size="1">
                <option value="0">________</option>
<%getStatesDropDown_Billing()%>
            	</select>
            </div>
            <div class="clear"></div>
        </div>
		<div class="Checkout_Field">Zip code <span class="Required">*required</span><br><input name="B_Zip" id="B_Zip" type="text" maxlength="5" class="rounded-glow w50" value="<%=strB_Zip%>"></div>
        <!-- END BILLING -->
            
        <!-- SHIPPING -->        
		<div class="Checkout_Header">Shipping Address</div>
        <div class="Checkout_Field" style="padding:5px; border:1px dashed #000;">
            <input type="checkbox" name="SameAsBilling" value="ON" style="border:0;"> Same as billing
        </div>
        <div class="Checkout_Field">
            First Name <span class="Required">*required</span><br><input name="S_FName" id="S_FName" type="text" class="rounded-glow w270" value="<%=strS_FName%>">
        </div>
        <div class="Checkout_Field">
            Last name <span class="Required">*required</span><br>
              <input name="S_LName" id="S_LName" type="text" class="rounded-glow w270" value="<%=strS_LName%>">
        </div>
        <div class="Checkout_Field">
            Address <span class="Required">*required</span><br><input name="S_Address" id="S_Address" type="text" class="rounded-glow w350" value="<%=strS_Address%>">
        </div>
        <div class="Checkout_Field">
            Address 2<br><input name="S_Address2" id="S_Address2" type="text" class="rounded-glow w350" value="<%=strS_Address2%>">
        </div>
        <div class="Checkout_Field">
            <div style="float:left;">
                City <span class="Required">*required</span><br><input name="S_City" id="S_City" type="text" class="rounded-glow w150" value="<%=strS_City%>">
            </div>
            <div style="float:left; margin-left:25px;">
                State <span class="Required">*required</span><br>
                <select name="S_State" id="S_State" class="rounded-glow" size="1">
                <option value="0">________</option>
<%getStatesDropDown_Shipping()%>
                </select>
            </div>
            <div class="clear"></div>
        </div>
        <div class="Checkout_Field">Zip code <span class="Required">*required</span><br><input name="S_Zip" id="S_Zip" type="text" maxlength="5" class="rounded-glow w50" value="<%=strS_Zip%>"></div>  
        <!-- END SHIPPING -->
        
        <!-- PAYMENT -->
        <div class="Checkout_Header">Credit Card Information</div>
        <div class="Checkout_Field">
            Type of Card <span class="Required">*required</span><br>
            <select name="CardType" id="CardType" class="rounded-glow" size="1">
                <option value="0">________</option>
                <option value="Visa" <% If strCardType = "Visa" Then Response.Write("selected")%>>Visa</option>
                <option value="MasterCard" <% If strCardType = "MasterCard" Then Response.Write("selected")%>>MasterCard</option>
                <option value="Discover" <% If strCardType = "Discover" Then Response.Write("selected")%>>Discover</option>
                <option value="Amex" <% If strCardType = "Amex" Then Response.Write("selected")%>>American Express</option>
            </select>
		</div>
        <div class="Checkout_Field">Card Number <span class="Required">*required</span><br><input name="CardNumber" id="CardNumber" type="text" class="rounded-glow w270" value="<%=strCardNumber%>"></div>
        <div class="Checkout_Field">
            Expiration Date <span class="Required">*required</span><br>
            MM <select name="ExpMo" id="ExpMo" size="1" class="rounded-glow w75">
                <option value="0">___</option>
                <option value="01" <% If strExpMo = "01" Then Response.Write("selected")%>>01</option>
                <option value="02" <% If strExpMo = "02" Then Response.Write("selected")%>>02</option>
                <option value="03" <% If strExpMo = "03" Then Response.Write("selected")%>>03</option>
                <option value="04" <% If strExpMo = "04" Then Response.Write("selected")%>>04</option>
                <option value="05" <% If strExpMo = "05" Then Response.Write("selected")%>>05</option>
                <option value="06" <% If strExpMo = "06" Then Response.Write("selected")%>>06</option>
                <option value="07" <% If strExpMo = "07" Then Response.Write("selected")%>>07</option>
                <option value="08" <% If strExpMo = "08" Then Response.Write("selected")%>>08</option>
                <option value="09" <% If strExpMo = "09" Then Response.Write("selected")%>>09</option>
                <option value="10" <% If strExpMo = "10" Then Response.Write("selected")%>>10</option>
                <option value="11" <% If strExpMo = "11" Then Response.Write("selected")%>>11</option>
                <option value="12" <% If strExpMo = "12" Then Response.Write("selected")%>>12</option>
                </select>
            YYYY <select name="ExpYear" id="ExpYear" size="1" class="rounded-glow w75">
                <option value="0">___</option>
                <option value="15" <% If strExpYear = "15" Then Response.Write("selected")%>>15</option>
                <option value="16" <% If strExpYear = "16" Then Response.Write("selected")%>>16</option>
                <option value="17" <% If strExpYear = "17" Then Response.Write("selected")%>>17</option>
                <option value="18" <% If strExpYear = "18" Then Response.Write("selected")%>>18</option>
                <option value="10" <% If strExpYear = "19" Then Response.Write("selected")%>>19</option>
                <option value="11" <% If strExpYear = "20" Then Response.Write("selected")%>>20</option>
                <option value="12" <% If strExpYear = "21" Then Response.Write("selected")%>>21</option>
                <option value="13" <% If strExpYear = "22" Then Response.Write("selected")%>>22</option>
                <option value="14" <% If strExpYear = "23" Then Response.Write("selected")%>>23</option>
                </select>
        </div>
        <div class="Checkout_Field">Card Verification Number <span class="Required">*required</span><br><input name="CardCVV" id="CardCVV" type="text" class="rounded-glow w50" maxlength="4" value="<%=strCVV%>"></div>
        <!-- END PAYMENT -->
        
        <!-- DISCOUNT -->
        <div class="Checkout_Header">Discount Code</div>
        <div class="Coupon_Code"><input type="text" name="CouponCode" class="rounded-glow w150" value="<%=strCouponCode%>"></div>
        <!-- END DISCOUNT -->
        
        <div class="Submit_Bar">
            <input type="button" id="mySubmit" class="large red button" name="Submit" value="Checkout">
        </div>
</form>
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
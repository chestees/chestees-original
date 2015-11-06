<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="/includes/globalLib.asp" -->
<!--#include virtual="/includes/adovbs.inc" -->
<%
'response.Write("C: " & cCustomerID)
'response.Write("V: " & cVisitorID)

If cCustomerID > 0 Then
	varBuyer = "Customer"
	cBuyerID = cCustomerID
	cVisitorID = 0
ElseIf cVisitorID > 0 Then
	varBuyer = "Visitor"
	cBuyerID = cVisitorID
	cCustomerID = 0
End If

OpenDB()

SQL = "SELECT FName, LName, Address, Address2, City, State, Zip, Email FROM tblBillingAddress WHERE Lock = 0 AND " & varBuyer & "ID = " & cBuyerID
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
	
SQL = "SELECT FName, LName, Address, Address2, City, State, Zip FROM tblShippingAddress WHERE Lock = 0 AND " & varBuyer & "ID = " & cBuyerID
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

CloseDB()
%>
<html>
<head>
<title>Secure Checkout</title>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="/css/style.css">
<link rel="shortcut icon" href="/images/favicon.ico" type="image/x-icon">
<script src="/js/jquery-1.4.2.min.js" type="text/javascript"></script>
<script src="/js/form.js" type="text/javascript"></script>
<script src="/js/form_Basic.js" type="text/javascript"></script>
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
<!-- End Header -->
<div id="Main">
	<div id="Main_Left">
    	<div style="padding-bottom:15px;">
        <a href="http://twitter.com/chestees" target="_blank"><img style="border:0;" src="/images/twitter-follow-me.png" alt="Follow us on Twitter" width="212" height="123" title="Follow us on Twitter"></a>
        </div>
        <div style="padding-bottom:15px;">
        <a href="/upload-photo/"><img src="/images/uploadPhoto.jpg" alt="Upload a Photo" title="Upload a Photo"></a>
        </div>
        <div style="padding-bottom:15px;">
        	<a href="http://www.damptshirts.com/"><img style="border:0;" src="/images/damp-vote.png" width="212" height="151" title="Vote for us on Damp T-Shirts" alt="Vote for us on Damp T-Shirts" /></a>
        </div>
    </div>
    <div id="Main_Basic"> <!-- START MAIN AREA CONTENT -->
        <div id="Main_Body">
            <div id="Main_Body_LeftColumn_Wide">
                <H1>Secure Checkout</H1>
            	
                <div id="myForm_ErrorRow" style="display:none; margin-top:10px; font-size:17px;"><div id="myForm_ErrorMessage"></div></div>
				<div id="myForm_Thinking" style="display:none; margin-top:10px; font-size:17px; text-align:center;">Submitting information...</div>
            
               	<div id="myForm">
<form id="htmlForm" action="/submit_Checkout.asp" method="post">
<input type="hidden" name="Redirect" value="/confirm/">
					<!-- BILLING -->
					<div style="float:left; padding:5px; border:1px solid #000; width:385px; margin-bottom:10px;">
                        <div style="float:left; width:375px; padding:5px; background-color:#CCCCCC; border-top:1px solid #b93636; border-bottom:1px solid #C84848; font-weight:bold;">
                            Billing Address
                        </div>
    
                        <div style="float:left; width:375px;">
                            <div style="float:left; width:200px; margin-top:10px;">
                                First Name*<br><input name="B_FName" type="text" id="input_1" style="width:150px" value="<%=strB_FName%>">
                            </div>
                            <div style="float:right; width:175px; margin-top:10px;">
                                Last Name*<br>
                                  <input name="B_LName" type="text" id="input_1" style="width:150px" value="<%=strB_LName%>">
                            </div>
                            
                            <div style="float:left; width:375px; margin-top:10px;">
                                Email*<br><input name="B_Email" type="text" id="input_1" style="width:250px" value="<%=strB_Email%>">
                            </div>
                            <div style="float:left; width:375px; margin-top:10px;">
                                Confirm Email*<br><input name="B_Email_Confirm" type="text" id="input_1" style="width:250px" value="<%=strB_Email%>">
                            </div>
                            
                            <div style="float:left; width:375px; margin-top:10px;">
                                Address*<br><input name="B_Address" type="text" id="input_1" style="width:250px" value="<%=strB_Address%>">
                            </div>
                            <div style="float:left; width:375px; margin-top:10px;">
                                Address 2<br><input name="B_Address2" type="text" id="input_1" style="width:250px" value="<%=strB_Address2%>">
                            </div>
                            <div style="float:left; width:200px; margin-top:10px;">
                                City*<br><input name="B_City" type="text" id="input_1" style="width:150px" value="<%=strB_City%>">
                            </div>
                            <div style="float:right; width:175px; margin-top:10px;">
                                State*<br>
                             	<select name="B_State" class="formSelect" style="width:150px" size="1">
                                <option value="0">________</option>
<%
OpenDB()

SQL = "SELECT State, Abbrev FROM tblState"
	Set rsStates = Conn.Execute(SQL)

	Do While Not rsStates.EOF
	
		strState = rsStates("State")
		strAbbrev = rsStates("Abbrev")
		
		Response.Write("<option value=" & strAbbrev)
		If strB_State = strAbbrev Then
			Response.Write(" selected")
		End If
		Response.Write(">" & strState & "</option>" & vbCRlf)
	
	rsStates.MoveNext
	Loop
	
	rsStates.Close
	Set rsStates = Nothing
	
CloseDB()
%>
                            	</select>
                            </div>
                            <div style="float:left; width:375px; margin-top:10px;">
                                Zip code*<br><input name="B_Zip" type="text" maxlength="5" id="input_1" style="width:50px" value="<%=strB_Zip%>">
                            </div>
                        </div>
					</div>
                    <!-- END BILLING -->
                        
                    <!-- SHIPPING -->
					<div style="float:left; padding:5px; border:1px solid #000; width:385px; margin-bottom:10px;">
                    
                        <div style="float:left; width:180px; height:20px; padding:5px; background-color:#CCCCCC; border-top:1px solid #b93636; border-bottom:1px solid #C84848; font-weight:bold;">
                            Shipping Address
                        </div>
                        <div style="float:right; width:185px; height:20px; padding:5px; background-color:#CCCCCC; border-top:1px solid #b93636; border-bottom:1px solid #C84848; text-align:right;">
                            <input type="checkbox" name="SameAsBilling" value="ON" style="border:0;"> Same as billing
                        </div>
                        
    
                        <div style="float:left; width:375px;">
                            <div style="float:left; width:200px; margin-top:10px;">
                                First Name*<br><input name="S_FName" type="text" id="input_1" style="width:150px" value="<%=strS_FName%>">
                            </div>
                            <div style="float:right; width:175px; margin-top:10px;">
                                Last name*<br>
                                  <input name="S_LName" type="text" id="input_1" style="width:150px" value="<%=strS_LName%>">
                            </div>
                            
                            <div style="float:left; width:375px; margin-top:10px;">
                                Address*<br><input name="S_Address" type="text" id="input_1" style="width:250px" value="<%=strS_Address%>">
                            </div>
                            <div style="float:left; width:375px; margin-top:10px;">
                                Address 2<br><input name="S_Address2" type="text" id="input_1" style="width:250px" value="<%=strS_Address2%>">
                            </div>
                            <div style="float:left; width:200px; margin-top:10px;">
                                City*<br><input name="S_City" type="text" id="input_1" style="width:150px" value="<%=strS_City%>">
                            </div>
                            <div style="float:right; width:175px; margin-top:10px;">
                                State*<br>
                             	<select name="S_State" class="formSelect" style="width:150px" size="1">
                                <option value="0">________</option>
<%
OpenDB()

SQL = "SELECT State, Abbrev FROM tblState"
	Set rsStates = Conn.Execute(SQL)

	Do While Not rsStates.EOF
	
		strState = rsStates("State")
		strAbbrev = rsStates("Abbrev")
		
		Response.Write("<option value=" & strAbbrev)
		If strS_State = strAbbrev Then
			Response.Write(" selected")
		End If
		Response.Write(">" & strState & "</option>")
	
	rsStates.MoveNext
	Loop
	
	rsStates.Close
	Set rsStates = Nothing
	
CloseDB()
%>
                            	</select>
                            </div>
                            <div style="float:left; width:375px; margin-top:10px;">
                                Zip code*<br><input name="S_Zip" type="text" maxlength="5" id="input_1" style="width:50px" value="<%=strS_Zip%>">
                            </div>
                        </div>    
					</div>
                    <!-- END SHIPPING -->
                    
                    <!-- PAYMENT -->
					<div style="float:left; padding:5px; border:1px solid #000; width:385px; margin-bottom:10px;">
                        <div style="float:left; width:375px; padding:5px; background-color:#CCCCCC; border-top:1px solid #b93636; border-bottom:1px solid #C84848; font-weight:bold;">
                            Credit Card Information
                        </div>
    
                        <div style="float:left; width:375px;">
                            <div style="float:left; width:200px; margin-top:10px;">
                                Type of Card*<br>
                                <select name="CardType" class="formSelect" style="width:150px" size="1">
                                    <option value="0">________</option>
                                    <option value="Visa" selected>Visa</option>
                                    <option value="MasterCard">MasterCard</option>
                                    <option value="Discover">Discover</option>
                                    <option value="Amex">American Express</option>
                                </select>
                            </div>
                            <div style="float:left; width:375px; margin-top:10px;">
                                Card Number*<br><input name="CardNumber" type="text" id="input_1" style="width:250px">
                            </div>
                            
                            <div style="float:left; width:375px; margin-top:10px;">
                                Expiration Date*<br>
                                MM <select name="ExpMo" size="1" class="formSelect" style="width:50px">
                                    <option value="0">___</option>
                                    <option value="01">01</option>
                                    <option value="02">02</option>
                                    <option value="03">03</option>
                                    <option value="04">04</option>
                                    <option value="05">05</option>
                                    <option value="06">06</option>
                                    <option value="07">07</option>
                                    <option value="08">08</option>
                                    <option value="09">09</option>
                                    <option value="10">10</option>
                                    <option value="11">11</option>
                                    <option value="12">12</option>
                                    </select>
                                YYYY <select name="ExpYear" size="1" class="formSelect" style="width:80px">
                                    <option value="0">___</option>
                                    <option value="08">08</option>
                                    <option value="09">09</option>
                                    <option value="10">10</option>
                                    <option value="11">11</option>
                                    <option value="12">12</option>
                                    <option value="13">13</option>
                                    <option value="14">14</option>
                                    <option value="15">15</option>
                                    <option value="16">16</option>
                                    <option value="17">17</option>
                                    <option value="18">18</option>
                                    </select>
                            </div>
                            <div style="float:left; width:375px; margin-top:10px;">
                                Card Verification Number*<br><input name="CardCVV" type="text" id="input_1" style="width:50px" maxlength="4">
                            </div>
                        </div>    
					</div>
                    <!-- END PAYMENT -->
                    
                    <!-- DISCOUNT -->
					<div style="float:left; padding:5px; border:1px solid #000; width:385px; margin-bottom:10px;">
                        <div style="float:left; width:375px; padding:5px; background-color:#CCCCCC; border-top:1px solid #b93636; border-bottom:1px solid #C84848; font-weight:bold;">
                            Discount Code
                        </div>
    
                        <div style="float:left; width:375px;">
                            <div style="float:left; width:200px; margin-top:10px;">
                                <input type="text" name="CouponCode" id="input_1" style="width:150px">
                            </div>
                        </div>    
					</div>
                    <!-- END DISCOUNT -->
                    
                    <div style="float:left; width:395px;">
                    	<input type="submit" id="mySubmit" name="Submit" value="Checkout">
                    </div>
</form>
                </div>
            </div>
            <div id="Main_Body_Random" style="font-size:12px;">
            	<div style="float:left; padding-bottom:10px; border-bottom:1px dashed #b93636; width:182px; text-align:center;">
                	<span id="siteseal"><script type="text/javascript" src="https://seal.godaddy.com/getSeal?sealID=lBM8wbUtZODPoQAPq2qArFpShU8Xya79V4QUwoJnAgxQWI4T3PGuEk"></script></span>
                </div>
                <div style="float:left; padding-bottom:10px; border-bottom:1px dashed #b93636; margin-top:10px; width:182px; text-align:center;">
                	<a href="https://www.paypal.com/us/verified/pal=jdiehl%40jasondiehl%2ecom" target="_blank"><img src="https://www.paypal.com/en_US/i/icon/verification_seal.gif" border="0" alt="Official PayPal Seal"></a>
                </div>
                <div style="float:left; padding-bottom:10px; margin-top:10px; border-bottom:1px dashed #b93636;">Orders arrive within 5-7 business days of the order being placed.</div>
                <div style="float:left; padding-bottom:10px; margin-top:10px; border-bottom:1px dashed #b93636;">All orders are shipped USPS Priority Mail&reg;</div>
                <div style="float:left; padding-bottom:10px; margin-top:10px;">Returns are accepted given the product is returned unwashed and unworn.</div>
            </div>
        </div>
    </div> <!-- END MAIN AREA CONTENT -->
</div> <!-- END MAIN AREA -->
<!--#include virtual="/incFooter.asp" -->
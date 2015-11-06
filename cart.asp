<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="/includes/globalLib.asp" -->
<!--#include virtual="/includes/adovbs.inc" -->
<%
If cCustomerID < 1 AND cVisitorID < 1 Then
	Response.Redirect("/enable-cookies/")
End If
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Shopping Cart</title>
<script src="/js/jquery-1.4.2.min.js" type="text/javascript"></script>
<script src="/js/jquery-ui-1.8.2.custom.min.js" type="text/javascript"></script>
<script src="/js/cart.js" type="text/javascript"></script>
<!--#include virtual="/incHeader.asp" -->
<div class="Main">
    <div class="TwoColumn_Left">
        <H1>Shopping Cart</H1>
    	<div id="Cart_Updated" class="Module"></div>
        <div class="Cart_Header">
            <div class="Cart_Col1">T-Shirt</div>
            <div class="Cart_Col2-3">Size</div>
            <div class="Cart_Col2-3">Price</div>
            <div class="Cart_Col4">Qty</div>
            <div class="clear"></div>
        </div>
        <form>
        <div id="myCart"></div>
        <div style="margin-top:10px;">
            <div style="float:left;"><input id="KeepShopping" type="button" class="medium blue button" value="<< keep shopping"></div>
            <div style="float:right;"><input id="myUpdate" type="button" class="medium blue button" value="update">&nbsp;&nbsp;<input id="myCheckout" type="button" class="medium red button" value="checkout >>"></div>
        	<div class="clear"></div>
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
<!--#include virtual="/incFooter.asp" -->
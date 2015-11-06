<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="/includes/globalLib.asp" -->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Chestees Funny T-Shirts<%=cKeywords_Title%></title>
<meta name="google-site-verification" content="VikoNTzWgwNboS3ua1rej9lg-nk43Tc6-wlEiaenan4" />
<meta property="og:title" content="Chestees T-Shirts" />
<meta property="og:description" content="<%=cDescription%>" />
<meta property="og:type" content="product" />
<meta property="og:site_name" content="Chestees T-Shirts" />
<meta property="fb:app_id" content="162492517129045"/>
<meta property="og:url" content="http://www.chestees.com/" />
<meta property="og:image" content="http://www.chestees.com/images/chestees_Logo.png" />
<script src="/js/jquery-1.4.2.min.js" type="text/javascript"></script>
<script src="/js/jquery-ui-1.8.2.custom.min.js" type="text/javascript"></script>
<script src="/js/email-signup.js" type="text/javascript"></script>
<script type="text/javascript">$(document).ready(function() {$('#chestees').addClass('ON')})</script>
<!--#include virtual="/incHeader.asp" -->
<div class="Main">
<!--#include virtual="/incBanners.asp" -->
    <div class="Main_Products"> <!-- START MAIN AREA CONTENT -->
<%
OpenDB()
'_____________________________________________________________________________________________
'CREATE THE PRODUCTS RECORDSET
SQL = "SELECT ProductID, Product, Image_Index FROM tblProduct WHERE Active = 1 AND Private = 0 AND CategoryID = 1 ORDER BY DisplayOrder"
	Set rsProducts = Conn.Execute(SQL)
	
	If Not rsProducts.EOF Then
		i = 0
		Do While Not rsProducts.EOF AND i < 4
			i = i+1

			intProductID = rsProducts("ProductID")
			strProduct = rsProducts("Product")
			strDirectory = Stripper(strProduct)
			strImage_Main = rsProducts("Image_Index")
			
			Response.Write("<a href=""funny-t-shirts/" & intProductID & "/" & strDirectory & "/""><img class=""Product Module"" src=""uploads/products/" & strImage_Main & """ alt=""" & strProduct & """ title=""" & strProduct & """></a>" & vbcrlf)

		rsProducts.MoveNext
		Loop
		
	End If
%>
<% Call EmailSignUp %>
<%
	If Not rsProducts.EOF Then
		i = 0
		Do While Not rsProducts.EOF
			i = i + 1
			If i=4 Then response.Write("</div><div style='float:right; width:636px; text-align:center; margin-bottom:10px;'><a target='_blank' href='http://www.damptshirts.com/?utm_campaign=Global&utm_source=Chestees&utm_medium=Homepage'><img border='0' src='images/damp-tshirts-banner.jpg'></a></div><div class='Main_Products'>")
			intProductID = rsProducts("ProductID")
			strProduct = rsProducts("Product")
			strDirectory = Stripper(strProduct)
			strImage_Main = rsProducts("Image_Index")
			
			Response.Write("<a href=""funny-t-shirts/" & intProductID & "/" & strDirectory & "/""><img class=""Product Module"" src=""uploads/products/" & strImage_Main & """ alt=""" & strProduct & """ title=""" & strProduct & """></a>" & vbcrlf)
			
		rsProducts.MoveNext
		Loop
		
	End If

rsProducts.Close
Set rsProducts = Nothing

CloseDB()
%>
    </div> <!-- END MAIN AREA CONTENT -->
    <div class="clear"></div>
</div> <!-- END MAIN AREA -->

<!--#include virtual="/incFooter.asp" -->
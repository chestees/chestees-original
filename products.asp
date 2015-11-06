<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" xmlns:fb="http://www.facebook.com/2008/fbml" xmlns:og="http://opengraphprotocol.org/schema/">
<!--#include virtual="/includes/globalLib.asp" -->
<!--#include virtual="/includes/adovbs.inc" -->
<%
intProductID = Request("ProductID")

'_____________________________________________________________________________________________
'OPEN DATABASE CONNECTION
Call OpenDB()

Set cmd = Server.CreateObject("ADODB.Command")
Conn.CursorLocation = 3
Set cmd.ActiveConnection = Conn
cmd.CommandText = "usp_ProductDetail"
cmd.CommandType = adCmdStoredProc

cmd.Parameters.Append cmd.CreateParameter("ProductID",adInteger,adParamInput)
cmd.Parameters("ProductID") = intProductID

Set rsProduct = cmd.Execute

strProduct = rsProduct("Product")
strDirectory = Stripper(strProduct)
strDescription = rsProduct("Description")
strImage_Main = rsProduct("Image_Index")
strImage_1 = rsProduct("Image_1")
strImage_2 = rsProduct("Image_2")
strImage_3 = rsProduct("Image_3")
strImage_4 = rsProduct("Image_4")
strDigg = rsProduct("Digg")
strStumbleTitle = rsProduct("StumbleTitle")
blnPrivate = rsProduct("Private")
blnSpecial = rsProduct("Special")
strTitle_Option = rsProduct("Title_Option")
strKeywords = rsProduct("Keywords")
blnActive = rsProduct("Active")

Set rsProduct = Nothing
Set cmd = nothing

'GET THE PRODUCT STYLES / OPTIONAL COLORS
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = Conn
cmd.CommandText = "usp_ProductStyle_NEW"
cmd.CommandType = adCmdStoredProc

cmd.Parameters.Append cmd.CreateParameter("ProductID",adInteger,adParamInput)
cmd.Parameters("ProductID") = intProductID
cmd.Parameters.Append cmd.CreateParameter("StyleCount",adInteger,adParamOutput)

Set rsStyle = cmd.Execute

curPrice = cmd.Parameters("StyleCount").Value
intStyleCount = cmd.Parameters("StyleCount").Value
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title><%=strTitle_Option%><%=cKeywords_Title%></title>

<meta property="og:title" content="<%=strProduct%> shirt from Chestees" />
<meta property="og:description" content="<%=cDescription%>" />
<meta property="og:type" content="product" />
<meta property="og:site_name" content="Chestees T-Shirts" />
<meta property="fb:app_id" content="162492517129045"/>
<meta property="og:url" content="http://www.chestees.com<%=request.servervariables("HTTP_X_ORIGINAL_URL")%>" />
<meta property="og:image" content="http://www.chestees.com/uploads/products/<%=strImage_Main%>" />

<script src="/js/jquery-1.4.2.min.js" type="text/javascript"></script>
<script src="/js/jquery-ui-1.8.2.custom.min.js" type="text/javascript"></script>
<script src="/video/flowplayer/flowplayer-3.1.4.min.js" type="text/javascript"></script>

<script src="/js/product.js" type="text/javascript"></script>
<script type="text/javascript" src="//assets.pinterest.com/js/pinit.js"></script>
<script type="text/javascript">
  (function() {
    var po = document.createElement('script'); po.type = 'text/javascript'; po.async = true;
    po.src = 'https://apis.google.com/js/plusone.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(po, s);
  })();
</script>
<!-- StumbleUpon Include -->
<script type="text/javascript">
  (function () {
    var li = document.createElement('script'); li.type = 'text/javascript'; li.async = true;
    li.src = ('https:' == document.location.protocol ? 'https:' : 'http:') + '//platform.stumbleupon.com/1/widgets.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(li, s);
  })();
</script>
<!--#include virtual="/incHeader.asp" -->
<div class="Main">
<!--#include virtual="/incBanners.asp" -->
    <div class="Content_Area Module"> <!-- START MAIN AREA CONTENT -->
    	<!-- START LEFT COLUMN AREA CONTENT -->
		<div class="MainProductsPage_L">
            <H1 class="Product_Title"><span><%=strProduct%></span><div class="Product_Price">$<%=rsStyle("Price")%></div><div class="clear"></div></H1>
            <div class="Product_Desc"><%=replace(strDescription,Chr(13),"<br />")%></div>
            <div class="clear"></div>
            <div class="Product_Detail_Mod"><img class="Product_IMG_Detail" id="Product_Big" src="/uploads/products/<%=strImage_1%>" /></div>
            <div class="Product_Thumbs">
                <%If Not IsNull(strImage_1) Then%>
                	<img class="Product_Thumb" id="Product_SM_1" src="/uploads/products/<%=strImage_1%>" />
                <%End If%>
                <%If Not IsNull(strImage_2) Then%>
                	<img class="Product_Thumb" id="Product_SM_2" src="/uploads/products/<%=strImage_2%>" />
                <%End If%>
                <%If Not IsNull(strImage_3) Then%>
                	<img class="Product_Thumb" id="Product_SM_3" src="/uploads/products/<%=strImage_3%>" />
                <%End If%>
                <%If Not IsNull(strImage_4) Then%>
                	<img class="Product_Thumb" id="Product_SM_4" src="/uploads/products/<%=strImage_4%>" style="margin:0;" />
                <%End If%>
                <div class="clear"></div>
            </div>
        </div>
        <!-- END LEFT COLUMN AREA CONTENT -->
		
        <!-- START RIGHT COLUMN AREA CONTENT -->
        <div class="MainProductsPage_R">
        <form>
        	<input type="hidden" name="ProductID" id="ProductID" value="<%=intProductID%>">   
            <img class="Product_Sub" src="/uploads/products/<%=strImage_Main%>">
            <div id="colorError"></div>
<%
If blnActive Then

	If Not rsStyle.EOF Then 
		If intStyleCount > 1 Then
%>
			<div style="margin-bottom:15px;">
				<select id="ProductStyleID" name="ProductStyleID" class="rounded-glow" style="width:180px" size="1">
                <option value="0">- select a color -</option>
<%
			Do While Not rsStyle.EOF
				curPrice = rsStyle("Price")
				intProductStyleID = rsStyle("ProductStyleID")
				strStyle = rsStyle("Style")
				Response.Write("<option value='" & intProductStyleID & "'>" & strStyle & "</option>")
			rsStyle.MoveNext
			Loop
%>
				</select>
            </div>
<%
		Else
			curPrice = rsStyle("Price")
			intProductStyleID = rsStyle("ProductStyleID")
			strStyle = rsStyle("Style")
			response.Write("<div class=""OnlyStyle"">Color: " & strStyle & "</div>")
			response.Write("<input type=""hidden"" value=""" & intProductStyleID & """ name=""ProductStyleID"">")
			response.Write("<div class=""clear""></div>")
		End If
	End If

Set rsStyle = Nothing
Set cmd = nothing
%>      
        <div style="margin-bottom:20px;">
<%
OpenDB()
'CREATE THE PRODUCTS SIZE RECORDSET
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = Conn
cmd.CommandText = "usp_ProductSize"
cmd.CommandType = adCmdStoredProc

cmd.Parameters.Append cmd.CreateParameter("ProductID",adInteger,adParamInput)
cmd.Parameters("ProductID") = intProductID
cmd.Parameters.Append cmd.CreateParameter("ProductStyleID",adInteger,adParamInput)
cmd.Parameters("ProductStyleID") = intProductStyleID

Set rsSize = cmd.Execute

If Not rsSize.EOF Then
	i=0
	Do While Not rsSize.EOF
		i=i+1
		intProductSizeID = rsSize("ProductSizeID")
		strSizeAbbr = rsSize("SizeAbbr")
		intQuantity = rsSize("Quantity")
		'Response.Write("QTY: " & intQuantity & "&nbsp;")
		If intQuantity > 0 Then			
			response.Write("<a class='Sizes")
				If i=1 Then Response.Write(" ON")
			response.Write("' id='" & intProductSizeID & "'>" & strSizeAbbr & "</a>")
		End If

	rsSize.MoveNext
	Loop

End If

Set rsSize = Nothing
set cmd = nothing
CloseDB()
%>
        </div>
        <div class="clear"></div>
        <div id="AddToCart" class="medium blue button" style="display:block; margin-bottom:20px;">Add to Cart</div>
<%Else%>
		<div class="Required" style="text-align:center; margin-bottom:20px;">This product is no longer available.</div>
<%End If%>
    </form>
	<div class="ShareProduct" style="text-align:left;">
      <div class="ShareWidgets">
        <div class="fb-like" data-href="http://www.chestees.com<%=request.servervariables("HTTP_X_ORIGINAL_URL")%>" data-send="false" data-layout="button_count" data-width="120" data-show-faces="true" data-font="arial"></div>
      </div>
      <div class="ShareWidgets">
      	<a target="_blank" href="http://pinterest.com/pin/create/button/?url=http://www.chestees.com<%=request.servervariables("HTTP_X_ORIGINAL_URL")%>&media=http://www.chestees.com/uploads/products/<%=strImage_1%>&description=<%=strTitle_Option%>" class="pin-it-button" count-layout="horizontal"><img border="0" src="//assets.pinterest.com/images/PinExt.png" title="Pin It" /></a>
      </div>
      <div class="ShareWidgets">
        <g:plusone size="medium" annotation="inline" width="120"></g:plusone>
      </div>
      <div class="ShareWidgets">
        <a href="https://twitter.com/share" class="twitter-share-button" data-count="horizontal">Tweet</a><script type="text/javascript" src="//platform.twitter.com/widgets.js"></script>
      </div>
      <div class="ShareWidgets">
        <script type="text/javascript" src="http://www.reddit.com/static/button/button1.js"></script>
      </div>
      <div class="ShareWidgets">
        <su:badge layout="1"></su:badge>
      </div>
    </div>
    <div id="sizeChart">Size chart</div>
    </div>
    <!-- END RIGHT COLUMN AREA CONTENT -->
    
	<div class="clear"></div>
    <!-- Comment Area -->
    	<div style="margin-top:10px;"><fb:comments href="www.chestees.com<%="/funny-t-shirts/" & intProductID & "/" & Stripper(strProduct) & "/"%>" num_posts="10" width="596"></fb:comments></div>
    </div>
    <!-- END Comment Area -->
    <div class="clear"></div>   
</div>
<!-- END MAIN AREA -->
<!--#include virtual="/incFooter.asp" -->
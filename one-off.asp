<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="/includes/globalLib.asp" -->
<!--#include virtual="/includes/adovbs.inc" -->
<%
'_____________________________________________________________________________________________
'OPEN DATABASE CONNECTION
Call OpenDB()

Set cmd = Server.CreateObject("ADODB.Command")
Conn.CursorLocation = 3
Set cmd.ActiveConnection = Conn
cmd.CommandText = "usp_ProductDetailSpecial"
cmd.CommandType = adCmdStoredProc

Set rsProduct = cmd.Execute

intProductID = rsProduct("ProductID")
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

intStyleCount = cmd.Parameters("StyleCount").Value
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title><%=strTitle_Option%><%=cKeywords_Title%></title>
<link rel="image_src" href="http://www.chestees.com/uploads/products/<%=strImage_Main%>" />
<meta name="title" content="<%=strProduct%> t-shirt from Chestees" />
<script src="/video/flowplayer/flowplayer-3.1.4.min.js" type="text/javascript"></script>
<script src="/js/jquery-1.4.2.min.js" type="text/javascript"></script>
<script src="/js/jquery-ui-1.8.2.custom.min.js" type="text/javascript"></script>
<script src="/js/product.js" type="text/javascript"></script>
<!--#include virtual="/incHeader.asp" -->
<div class="Main">
<!--#include virtual="/incBanners.asp" -->
<div class="Content_Area Module">
  <div style="float:left; width:187px;"><img src="/images/vault.jpg" /></div>
  <div style="float:right; width:409px; font-size:11px;">Welcome to the vault! The vault has shirts you can't buy on any given day. This is where you can find your favorite design, but on a different color t-shirt than we sell normally. Why is this? Well we like to experiment and why waste a perfectly good shirt? We figure we'll sell it at a super discount. Check back often because this page changes often.</div>
  <div class="clear"></div>
</div>
<div class="Content_Area Module">
  <!-- START MAIN AREA CONTENT -->
  <!-- START LEFT COLUMN AREA CONTENT -->
  <div class="MainProductsPage_L">
    <div class="Product_Title">
      <%If len(strProduct) > 21 Then%>
      <H1 class="Small_Title"><%=strProduct%></H1>
      <%Else%>
      <H1 class="Large_Title"><%=strProduct%></H1>
      <%End If%>
    </div>
    <div class="Product_Price">$<%=rsStyle("Price")%></div>
    <div class="clear"></div>
    <!-- SOCIAL BOOKMARK -->
    <div class="ShareProduct">
      <div id="ShareIcon"><a class="Social Facebook" href="http://www.facebook.com/share.php?u=<%= Server.URLEncode("http://www.chestees.com" & request.servervariables("HTTP_X_ORIGINAL_URL") & "from,chestees_facebook") %>&t=<%=strProduct%> from Chestees" rel="nofollow external"></a></div>
      <div id="ShareIcon"><a class="Social Twitter" href="http://twitter.com/home/?status=<%= Server.URLEncode("Look what I found at @Chestees: " & "http://www.chestees.com" & request.servervariables("HTTP_X_ORIGINAL_URL") & "from,chestees_twitter")%>" rel="nofollow external"></a></div>
      <div id="ShareIcon"><a class="Social Myspace" href="http://www.myspace.com/Modules/PostTo/Pages/?u=<%= Server.URLEncode("http://www.chestees.com" & request.servervariables("HTTP_X_ORIGINAL_URL") & "from,chestees_myspace") %>&l=3&c=" rel="nofollow external"></a></div>
      <div id="ShareIcon"><a class="Social Digg" href="http://digg.com/submit??phase=2&url=<%= Server.URLEncode("http://www.chestees.com" & request.servervariables("HTTP_X_ORIGINAL_URL") & "from,chestees_digg") %>&title=<%= Server.URLEncode(strProduct) %>" rel="nofollow external"></a></div>
      <div id="ShareIcon"><a class="Social Stumble" href="http://www.stumbleupon.com/submit?&url=<%= Server.URLEncode("http://www.chestees.com" & request.servervariables("HTTP_X_ORIGINAL_URL") & "from,chestees_stumbleupon") %>&title=<%= Server.URLEncode(strProduct) %>" rel="nofollow external"></a></div>
      <div id="ShareIcon"><a class="Social Delicious" href="http://del.icio.us/post?&url=<%= Server.URLEncode("http://www.chestees.com" & request.servervariables("HTTP_X_ORIGINAL_URL") & "from,chestees_delicious") %>&title=<%= Server.URLEncode(strProduct) %>" rel="nofollow external"></a></div>
      <div class="clear"></div>
    </div>
    <!-- SOCIAL BOOKMARK -->
    <div class="Product_Desc"><%=replace(strDescription,Chr(13),"<br />")%></div>
    <div class="clear"></div>
    <%
If Not rsStyle.EOF Then
%>
    <div style="margin-top:10px;" id="AvailableStyles"><b>Available in these styles:</b></div>
    <%
	Do While Not rsStyle.EOF
	
		curPrice = rsStyle("Price")
		strStyle = rsStyle("Style")
%>
    <div class="AvailableIn"><%=strStyle%> - <%=FormatCurrency(curPrice,0)%></div>
    <%
	rsStyle.MoveNext
	Loop
	rsStyle.MoveFirst
End If
%>
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
      <%
If Not rsStyle.EOF Then 

	If intStyleCount > 1 Then
%>
      <div style="margin-bottom:15px;">
        <select id="ProductStyleID" name="ProductStyleID" class="rounded-glow" style="width:180px" size="1">
          <%
		Do While Not rsStyle.EOF
			curPrice = rsStyle("Price")
			intProductStyleID = rsStyle("ProductStyleID")
			strStyle = rsStyle("Style")
			Response.Write("<option value=""" & intProductStyleID & """>" & strStyle & "</option>")
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
		response.Write("<div class=""OnlyStyle"">" & strStyle & ": $" & curPrice & "</div>")
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
	Do While Not rsSize.EOF
		intProductSizeID = rsSize("ProductSizeID")
		strSizeAbbr = rsSize("SizeAbbr")
		intQuantity = rsSize("Quantity")
		'Response.Write("QTY: " & intQuantity & "&nbsp;")
		If intQuantity > 0 Then			
			response.Write("<a class=""Sizes"" id=""" & intProductSizeID & """>" & strSizeAbbr & "</a>")
		End If

	rsSize.MoveNext
	Loop

End If

Set rsSize = Nothing
set cmd = nothing
CloseDB()
%>
      </div>
      <div style="margin:30px 0;"><a class="small gray button" href="javascript:popUpWindow('/sizing-info/','40','20','520','425');">Size chart</a></div>
    </form>
  </div>
  <!-- END RIGHT COLUMN AREA CONTENT -->
  <div class="clear"></div>
</div>
<!-- END MAIN AREA -->

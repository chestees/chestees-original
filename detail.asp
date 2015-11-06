<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="includes/globalLib.asp" -->
<!--#include virtual="includes/adovbs.inc" -->
<%
intDiggID = Request("i")
	
'_____________________________________________________________________________________________
'OPEN DATABASE CONNECTION
Call OpenDB()

Set cmd = Server.CreateObject("ADODB.Command")
Conn.CursorLocation = 3
Set cmd.ActiveConnection = Conn
cmd.CommandText = "usp_Digg_AffiliateLink"
cmd.Parameters.Append cmd.CreateParameter("DiggID",adInteger,adParamInput)
cmd.Parameters("DiggID") = intDiggID
cmd.CommandType = adCmdStoredProc
Set rsDiggLink = cmd.Execute

strLink_Full = rsDiggLink("Link")
strLinkPrefix = rsDiggLink("LinkPrefix")
strLinkSuffix = rsDiggLink("LinkSuffix")
'If instr(strLink_Full,"http://") AND strLinkPrefix <> ""  AND instr(strLink_Full,"shareasale") Then
If strLinkPrefix <> "" Then
	If instr(strLinkPrefix,"shareasale") > 0 Then
		strLinkPrefix = replace(rsDiggLink("LinkPrefix"),"afftrack=","afftrack="&intDiggID)
		strLink = strLinkPrefix & replace(strLink_Full,"http://","")
	Else
		strLink = strLinkPrefix & strLink_Full
	End If
Else
	strLink = strLink_Full & strLinkSuffix
End If
strTitle = rsDiggLink("Title")
strImage = rsDiggLink("Image")
strDiggStore = rsDiggLink("DiggStore")
CloseDB()
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<meta name="robots" content="index, follow" />
<meta property="og:title" content="<%=strTitle%> t-shirt from <%=strDiggStore%>" />
<meta property="og:description" content="Check it out!" />
<meta property="og:type" content="product" />
<meta property="og:url" content="http://www.chestees.com<%=request.servervariables("HTTP_X_ORIGINAL_URL")%>" />
<meta name="keywords" content="<%If strKeywords <> "" Then Response.Write(strKeywords & " ")%><%=cKeywords%>" />
<meta property="og:image" content="<%=strImage%>" />
<script src="js/jquery-1.4.2.min.js" type="text/javascript"></script>
<script src="js/jquery-ui-1.8.2.custom.min.js" type="text/javascript"></script>
<title><%=strTitle%> from <%=strDiggStore%></title>
<script>
  window.fbAsyncInit = function () {
    FB.init({
      appId: '230800510267354', // App ID
      channelUrl: '//www.chestees.com/channel.html', // Channel File
      status: true, // check login status
      cookie: true, // enable cookies to allow the server to access the session
      xfbml: true  // parse XFBML
    });

    _ga.trackFacebook(); //Google Analytics tracking
  };

  // Load the SDK Asynchronously
  (function (d) {
    var js, id = 'facebook-jssdk', ref = d.getElementsByTagName('script')[0];
    if (d.getElementById(id)) { return; }
    js = d.createElement('script'); js.id = id; js.async = true;
    js.src = "//connect.facebook.net/en_US/all.js";
    ref.parentNode.insertBefore(js, ref);
  } (document));
</script>
<!--#include virtual="/incHeader.asp" -->
<div class="Main clearfix">
	<!--#include virtual="/incBanners.asp" -->
  <div class="Content_Area Module">    
    <h1><%=strTitle%> from <%=strDiggStore%></h1>
    <div style="float:left; width:500px;">
		  <a href="<%=strLink %>" target="_blank"><img src="<%=strImage%>" class="detail_img" /></a>
      <a href="<%=strLink %>" target="_blank" class="detail_link">Get the "<%=strTitle%>" t-shirt from <%=strDiggStore%></a>
    </div>
    <div style="float:right;">
      <!-- SOCIAL BOOKMARK -->
      <div class="share clearfix">
        <div class="share_widgets facebook">
          <script type="text/javascript">_ga.trackFacebook();</script>
          <div class="fb-like" data-send="false" data-layout="box_count" data-show-faces="false"></div>
        </div>
        <div class="share_widgets google">
          <g:plusone size="tall"></g:plusone>
        </div>
        <div class="share_widgets twitter">
          <a href="https://twitter.com/share" class="twitter-share-button" data-text="Check out this t-shirt" data-via="chestees" data-hashtags="funnyshirts" data-count="vertical">Tweet</a>
          <script>!function (d, s, id) { var js, fjs = d.getElementsByTagName(s)[0]; if (!d.getElementById(id)) { js = d.createElement(s); js.id = id; js.src = "//platform.twitter.com/widgets.js"; fjs.parentNode.insertBefore(js, fjs); } } (document, "script", "twitter-wjs");</script>  
        </div>
        <div class="share_widgets pinterest">
          <a href="http://pinterest.com/pin/create/button/?url=<%=Server.URLEncode("http://www.chestees.com" & request.servervariables("HTTP_X_ORIGINAL_URL"))%>&media=<%= strImage %>" class="pin-it-button" count-layout="vertical">Pin It</a>
          <script type="text/javascript" src="http://assets.pinterest.com/js/pinit.js"></script>
        </div>
        <div class="share_widgets reddit">
          <script type="text/javascript" src="http://www.reddit.com/static/button/button3.js"></script>
        </div>
      </div>
    	<!-- SOCIAL BOOKMARK -->
    </div>
    <div class="clear"></div>
	</div> 
</div> 
</body>
</html>

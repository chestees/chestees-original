<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="/includes/globalLib.asp" -->
<!--#include virtual="/includes/adovbs.inc" -->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Uneetee T-Shirt Designs<%=cKeywords_Title%></title>
<script src="/js/jquery-1.4.2.min.js" type="text/javascript"></script>
<script src="/js/jquery-ui-1.8.2.custom.min.js" type="text/javascript"></script>
<script src="/js/email-signup.js" type="text/javascript"></script>
<script type="text/javascript">$(document).ready(function() {$('#uneetee').addClass('ON')})</script>
<!--#include virtual="/incHeader.asp" -->
<div class="Main clearfix">
<!--#include virtual="/incBanners.asp" -->
    <div class="Main_Products"> <!-- START MAIN AREA CONTENT -->
    	<div class="AffiliateHeader"><H1>T-Shirt Designs by Uneetee</H1></div>
        <div class="CouponArea Module">
            <div class="Left"></div>
            <div class="Right"><a href="http://www.shareasale.com/r.cfm?u=323844&b=107913&m=12558&afftrack=CT_UNEETEE&urllink=www.uneetee.com">VISIT UNEETEES</a></div>
            <div class="clear"></div>
        </div>
<%
OpenDB()
'_____________________________________________________________________________________________
'CREATE THE PRODUCTS RECORDSET
Set cmd = Server.CreateObject("ADODB.Command")
Conn.CursorLocation = 3
Set cmd.ActiveConnection = Conn

cmd.CommandText = "usp_Digg_ListFromTag"

cmd.Parameters.Append cmd.CreateParameter("TagID",adInteger,adParamInput)
cmd.Parameters("TagID") = 111

cmd.CommandType = adCmdStoredProc
Set rsTees = cmd.Execute
	
	If Not rsTees.EOF Then
		i = 0
		Do While Not rsTees.EOF
			i = i+1
			
			If i = 5 Then Call EmailSignUp
			
			If i=8 Then response.Write("</div><div style='float:right; width:636px; text-align:center; margin-bottom:10px;'><a target='_blank' href='http://www.damptshirts.com/?utm_campaign=Global&utm_source=Chestees&utm_medium=HomeUneetee'><img border='0' src='http://www.chestees.com/images/damp-tshirts-banner.jpg'></a></div><div class='Main_Products'>")
			
			intDiggID = rsTees("DiggID")
			strImage = rsTees("Image")
			strTitle = rsTees("Title")
			strLink_Full = rsTees("Link")
			strLinkPrefix = replace(rsTees("LinkPrefix"),"afftrack=","afftrack="&intDiggID)
			strLinkSuffix = rsTees("LinkSuffix")
			If instr(strLink_Full,"http://") AND strLinkPrefix <> "" Then
				strLink = strLinkPrefix & replace(strLink_Full,"http://","")
			Else
				strLink = strLink_Full & strLinkSuffix
			End If
			
			Response.Write("<a class='Affiliate' href='/t-shirts/detail/" & intDiggID & "/" & Stripper(strTitle) & "/'><img class='Product Module' src='" & strImage & "' alt='" & strTitle & "' title='" & strTitle & "' border='0'></a>")
		
		rsTees.MoveNext
		Loop
		
	End If

rsTees.Close
Set rsTees = Nothing

CloseDB()
%>
    </div> <!-- END MAIN AREA CONTENT -->
</div> <!-- END MAIN AREA -->
<!--#include virtual="/incFooter.asp" -->
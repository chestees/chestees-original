<!--#include virtual="/includes/globalLib.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Administration</title>
<script src="http://code.jquery.com/jquery-1.4.2.min.js" type="text/javascript" /></script>
<script type="text/javascript">
$(function() {
						 
	id = "http://www.chestees.com/funny-t-shirts/3/come-on-the-kick-drum/";
	id = id + ",http://www.chestees.com/funny-t-shirts/21/merry-christmas-kiss-my-ass/";
	id = id + ",http://www.chestees.com/funny-t-shirts/76/sgt-dr-pepper/";
	id = id + ",http://www.chestees.com/funny-t-shirts/75/national-grammar-rodeo/";
	id = id + ",http://www.chestees.com/funny-t-shirts/24/ill-most-likely-kill-you/";
	id = id + ",http://www.chestees.com/funny-t-shirts/63/green-is-good/";
	id = id + ",http://www.chestees.com/funny-t-shirts/62/im-taking-my-talents-to-south-beach/";
	id = id + ",http://www.chestees.com/funny-t-shirts/70/american-handegg/";
	id = id + ",http://www.chestees.com/funny-t-shirts/66/go-ahead-touch-my-junk/";
	id = id + ",http://www.chestees.com/funny-t-shirts/71/youve-been-warned/";
	id = id + ",http://www.chestees.com/funny-t-shirts/63/green-is-good/";
	id = id + ",http://www.chestees.com/funny-t-shirts/59/handsome-mens-club/";
	id = id + ",http://www.chestees.com/funny-t-shirts/56/bacon-is-good-for-me/";
	id = id + ",http://www.chestees.com/funny-t-shirts/46/deluxe-hugs/";
	id = id + ",http://www.chestees.com/funny-t-shirts/45/brett-favre-from-over/";
	id = id + ",http://www.chestees.com/funny-t-shirts/41/oscar-trophy/";
	id = id + ",http://www.chestees.com/funny-t-shirts/61/betty-white-is-a-friend/";
	id = id + ",http://www.chestees.com/funny-t-shirts/40/a-wolfpack-of-one/";
	id = id + ",http://www.chestees.com/funny-t-shirts/13/talk-to-the-han/";
	id = id + ",http://www.chestees.com/funny-t-shirts/28/potato-chip-pirate/";
	id = id + ",http://www.chestees.com/funny-t-shirts/2/a-very-uncomfortable-place/";
	id = id + ",http://www.chestees.com/funny-t-shirts/11/a-plethora-of-pinatas/";
	id = id + ",http://www.chestees.com/funny-t-shirts/35/the-banana-shack/";
	id = id + ",http://www.chestees.com/funny-t-shirts/29/nerdo-loco/";
	id = id + ",http://www.chestees.com/funny-t-shirts/19/sturdy-wings/";
	id = id + ",http://www.chestees.com/funny-t-shirts/20/take-me-to-the-volcano/";
	id = id + ",http://www.chestees.com/funny-t-shirts/17/dancing-leads-to-sex/";
	id = id + ",http://www.chestees.com/funny-t-shirts/18/dead-sexy/";
	id = id + ",http://www.chestees.com/funny-t-shirts/22/goats-do-it-for-the-kids/";
	id = id + ",http://www.chestees.com/funny-t-shirts/26/nakatomi-plaza-security/";
	id = id + ",http://www.chestees.com/funny-t-shirts/25/never-forget/";
	id = id + ",http://www.chestees.com/funny-t-shirts/5/cheese-and-rice/";
	id = id + ",http://www.chestees.com/funny-t-shirts/12/funny-how/";

	url = "http://graph.facebook.com/?ids="+id+"&callback=?";
	$.getJSON(url, function(data) {
		//alert("D: " + data);
		$.each(data, function(i,item){
			$("#mytext").append("<br /><br />URL: " + item.id);
			$("#mytext").append("<br />Shares: " + item.shares);
			$("#mytext").append("<br />Comments: " + item.comments);
		});
	});
});
</script>
<title><%=cFriendlySiteName%> | Administration</title>
<link rel="stylesheet" href="/css/stylesheet.css">
<link rel="shortcut icon" href="/images/favicon.ico" type="image/x-icon">
</head>

<body leftmargin="0" topmargin="0" marginWidth="0" marginHeight="0">
<table width="804" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
  	<td bgcolor="#A13846" width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
	<td bgcolor="#EEF2FC">
	  <table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
	  	<tr>
		  <td colspan="3"><img src="<%=cAdminPath%>images/header.jpg"></td>
		</tr>
		<tr>
		  <td width="153" valign="top">
		  	<table border="0" cellspacing="0" cellpadding="0">
			  <tr>
			  	<td align="center">
				  <table border="0" cellspacing="0" cellpadding="0">
				  	<tr>
				  	  <td width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="38"></td>
					  <td width="151" class="NavTitle">Navigation</td>
					</tr>
				  </table>
				</td>
			  </tr>
			  <tr>
			  	<td>
<!--#include virtual="/incNav.asp" -->
				</td>
			  </tr>
			</table>
		  </td>
		  <td width="2" bgcolor="#A13846"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
		  <td width="645" valign="top">
		  	<table border="0" cellspacing="0" cellpadding="0" width="645">
			  <tr>
			  	<td>
				  <table border="0" cellspacing="0" cellpadding="0">
				  	<tr>
				  	  <td width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="38"></td>
					  <td width="621" class="PageTitle" align="right">FACEBOOK</td>
					  <td width="20" bgcolor="#A13846"><img src="<%=cAdminPath%>images/filler.gif" width="20" height="1"></td>
					  <td width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td colspan="4"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="10"></td>
					</tr>
				  </table>
				</td>
			  </tr>
			  <tr>
			  	<td>
				  <table border="0" cellspacing="0" cellpadding="4" width="599" align="center">
					<tr><td><div id="mytext"></div></td></tr>
					<tr>
					  <td style="border-top:1px dashed #A13846;"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="10"></td>
					</tr>
				  </table>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
	<td bgcolor="#A13846" width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
  </tr>
  <tr>
  	<td colspan="3" bgcolor="#A13846" height="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="2"></td>
  </tr>
</table>
<!--#include virtual="/incFooter.asp" -->

</body>
</html>
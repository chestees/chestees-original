<%
Dim Conn
Dim strConn

cServerName = Request.ServerVariables("Server_Name")
CONST cSiteName = "Chestees"
'cFriendlySiteName = "Chestees"
'cURL = "chestees.com"
CONST cSiteURL = "http://www.chestees.com"
CONST cColor = "#E2E6EF"
CONST cKeywords_Title = " - funny t shirts - funny t-shirts - funny tshirts - funny tees - screen printed shirts - busted tees - snorg tees - deez teez - uneetees"
CONST cDescription = "Chestees are hand screen-printed funny t-shirts that will make you laugh until you stop."
CONST ET_User = ""
CONST ET_Password = ""

CONST ListId = 2319

CONST strAPIUrl = "https://api.s4.exacttarget.com/integrate.aspx"

cKeywords = "funny t-shirts, "
cKeywords = cKeywords & "funny t shirts, "
cKeywords = cKeywords & "funny tshirts, "
cKeywords = cKeywords & "t-shirts, "
cKeywords = cKeywords & "t shirts, "
cKeywords = cKeywords & "tshirts, "
cKeywords = cKeywords & "funny tees, "
cKeywords = cKeywords & "screen printed tees, "
cKeywords = cKeywords & "screen printed t-shirts, "
cKeywords = cKeywords & "screen printed tshirts, "
cKeywords = cKeywords & "alternative apparel, "
cKeywords = cKeywords & "busted tees, "
cKeywords = cKeywords & "snorg tees, "
cKeywords = cKeywords & "deez teez, "
cKeywords = cKeywords & "uneetee shirts,"
cKeywords = cKeywords & "snorgtees, snorgshirts, snorg,"

cCustomerID = Request.Cookies(cSiteName)("CustomerID")
cVisitorID = Request.Cookies(cSiteName)("VisitorID")
blnPrivate = Request.Cookies(cSiteName)("Private")
'response.Write("C: " & cCustomerID)
'response.Write("V: " & cVisitorID)

cShippingCost = 5
cFreeShippingNum = 100000
cMailHost = "mail.chestees.com"

strConn = "Driver={SQL Server};Server=;Database=;Uid=;Pwd=;"

'_________________________________________________________________________________________________
'Place this subroutine to open db connection
Private Sub OpenDB()	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.ConnectionString = strConn
	Conn.Open
End Sub
'_________________________________________________________________________________________________
'place this subroutine to close db connection
Private Sub CloseDB()
	Conn.Close
	Set Conn = Nothing
End Sub

Response.Expires 	= 0
Response.Buffer 	= True

Function Header()
	intCartCount = 0

	If cVisitorID > 0 Then
		Call OpenDB()
		'_____________________________________________________________________________________________
		'GET CART ITEMS COUNT
		SQL = "SELECT Quantity FROM tblCart WHERE VisitorID = " & cVisitorID & " AND Purchased = 0"
		Set rsCart = Server.CreateObject("ADODB.Recordset") 
		rsCart.Open SQL,Conn,3,3
		'intCartCount = rsCart.recordcount
		If Not rsCart.EOF Then
			Do While Not rsCart.EOF
				intCartCount = intCartCount + rsCart("Quantity")
			rsCart.MoveNext
			Loop
		End If
		rsCart.Close
		Set rsCart = Nothing

		strString = strString & "<li><a class='myCart' href='" & cSiteURL & "cart/'>my cart (<span id='CartCount'>" & intCartCount & "</span>)</a></li>"
		
		Call CloseDB()
	End If
	
	strHeader = "<div class='Header'>"
	strHeader = strHeader & "<div class='Nav'>"
	strHeader = strHeader & "<ul>"
	strHeader = strHeader & "<li><a id='chestees' href='" & cSiteURL & "/'>Chestees</a></li>"
	'strHeader = strHeader & "<li><a id='snorgtees' href='" & cSiteURL & "/snorg-tees/'>Snorg Tees</a></li>"
	'strHeader = strHeader & "<li><a id='bustedtees'a href='" & cSiteURL & "/busted-tees/'>Busted Tees</a></li>"
	'strHeader = strHeader & "<li><a id='deezteez' href='" & cSiteURL & "/deez-teez/'>Deez Teez</a></li>"
	'strHeader = strHeader & "<li><a id='uneetee' href='" & cSiteURL & "/uneetee/'>Uneetees</a></li>"
	strHeader = strHeader & "</ul>"
	strHeader = strHeader & "</div>"
	strHeader = strHeader & "<a class='Logo' href='/'><img src='/images/chestees_Logo.png' alt='Chestees Funny T-Shirts' width='239' height='66' title='Chestees Funny T-Shirts'></a>"
	strHeader = strHeader & "<div class=""clear""></div>"
	strHeader = strHeader & "</div>"
	strHeader = strHeader & "<div class='SubHeader Module'>"
	strHeader = strHeader & "<div class='SubHeaderLeft'>"
	strHeader = strHeader & "<div class='SubNav'>"
	strHeader = strHeader & "<ul>"
	strHeader = strHeader & "<li><a href='http://funny-tshirts-blog.chestees.com/'>blog</a></li>"
	strHeader = strHeader & "<li><a href='" & cSiteURL & "/contact-chestees/'>contact</a></li>"
	strHeader = strHeader & "<li><a href='" & cSiteURL & "/faq-chestees/'>faq</a></li>"
	strHeader = strHeader & "<li><a href='" & cSiteURL & "/about-chestees/'>about us</a></li>"
	strHeader = strHeader & strString
	strHeader = strHeader & "</ul>"
	strHeader = strHeader & "</div>"
	strHeader = strHeader & "</div>"
	strHeader = strHeader & "<div class='HeaderParagraph'>T-shirts by Chestees are funny original screen-printed t-shirts.</div>"
	strHeader = strHeader & "</div>"
	
	Response.Write(strHeader)
End Function

'*******************************************************
'BEGIN STRIPPER
Function Stripper(inFileName)
	dim sOut,strorigFileName,arrSpecialChar,intCounter
	arrSpecialChar  = Array("%20","%"," ","#","+","(",")","&","$","@","!","*","<",">","?","/","|","\",",","'",":","...","�",".")
	strorigFileName = replace(inFileName," ","-")
	intCounter = 0
	Do Until intCounter = 24
		sOut = replace(strorigFileName,arrSpecialChar(intCounter),"")
	  	intCounter = intCounter + 1
	  	strorigFileName = lcase(sOut)
	 Loop
	 'StripSpecialChar = response.write(strorigFileName)
	 Stripper = replace(strorigFileName,"---","-")
End Function
'END STRIPPER
'*******************************************************

Function EmailSignUp()
	strEmailString = "<div class='EmailSignUp Module'>"_          
		&"<div id='EmailSignUp_Title'>Sign up for Instant Discounts and News</div>"_
		&"<div id='EmailSignUp_Input'>"_
			&"<div style='float:left;'><input type='text' name='Email' value='' id='txtEmailSignUp' maxlength='75'></div>"_
			&"<div style='float:right;' id='btnSignUp'><a class='signup gray button'>Sign Up</a></div>"_
			&"<div class='clear'></div>"_
		&"</div>"_
		&"<div id='EmailSignUp_Text'>At Chestees, we love to keep in touch with our friends. We have strict 'no spam' policy and you can opt out at any time. We normally email only to announce new products and offers. If you don't want to give out your email, then become our friend on <a target='_blank' href='http://www.facebook.com/chestees'>Facebook</a>.</div>"_
	&"</div>"
	Response.Write(strEmailString)
End Function

Function varBuyer()
	If cCustomerID > 0 Then
		varBuyer = "Customer"
	ElseIf cVisitorID > 0 Then
		varBuyer = "Visitor"
	End If
End Function

Function cBuyerID()
	If cCustomerID > 0 Then
		cBuyerID = cCustomerID
	ElseIf cVisitorID > 0 Then
		cBuyerID = cVisitorID
	End If
End Function

Function getStatesDropDown_Billing()
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
End Function
Function getStatesDropDown_Shipping()
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
End Function

Function SQLEncode(strValue)
	strValue = Replace(strValue,"'","''")
	If strValue = "" then
		SQLEncode = "NULL"
	Else 
		SQLEncode = "'" & strValue & "'"
	End If	
End Function

Function SQLDateEncode(strValue)
	strValue = Replace(strValue,"'","''")
	
	If strValue = "" then
		SQLDateEncode = "NULL"
	Else 
		SQLDateEncode = "'" & strValue & "'"
	End If
End Function

Function SQLNumEncode(strValue)
	If strValue = "" OR Not IsNumeric(strValue) then
		SQLNumEncode = "NULL"
	Else 
		SQLNumEncode = strValue
	End If
End Function

Function CheckBoxValue(i_value)
	If Lcase(i_value) = "on" then
		CheckBoxValue = 1
	Else
		CheckBoxValue = 0
	End If
End Function

'_____________________________________________________________________________________________
'CREATE ACCOUNT EMAIL
Function Email_SignUp()
		
	TextBody = "" & vbcrlf & _
	"<html>" & vbcrlf & _
	"<head>" & vbcrlf & _
	"<title>chestees.com</title>" & vbcrlf & _
	"<style type=text/css>" & vbcrlf & _
	"<!--" & vbcrlf & _
	  "body {" & vbcrlf & _
    	"margin: 0px;" & vbcrlf & _
		"padding: 0px;" & vbcrlf & _
		"background-image:url(http://www.chestees.com/images/bgIMG.jpg);" & vbcrlf & _
    	"background-color: #2f1111;" & vbcrlf & _
	  "}" & vbcrlf & _
	  "td {" & vbcrlf & _
	  	"font-family: Georgia, Tahoma, verdana;" & vbcrlf & _
		"font-size: 14px;" & vbcrlf & _
		"color: #000000;" & vbcrlf & _
	  "}" & vbcrlf & _
	"-->" & vbcrlf & _
	"</style>" & vbcrlf & _
	"</head>" & vbcrlf & _
	"<body leftmargin=0 topmargin=0 marginWidth=0 marginHeight=0>" & vbcrlf & _
	  "<table width=582 border=0 cellspacing=0 cellpadding=0 align=center bgcolor=#FFFFFF>" & vbcrlf & _
  		"<tr>" & vbcrlf & _
  		  "<td width=1 bgcolor=#000000><img src=http://www.chestees.com/images/filler.gif width=1></td>" & vbcrlf & _
 	 	  "<td width=580><img src=http://www.chestees.com/images/header_CreateAccount.jpg></td>" & vbcrlf & _
		  "<td width=1 bgcolor=#000000><img src=http://www.chestees.com/images/filler.gif width=1></td>" & vbcrlf & _
  		"</tr>" & vbcrlf & _
	  	"<tr>" & vbcrlf & _
  		  "<td width=1 bgcolor=#000000><img src=http://www.chestees.com/images/filler.gif width=1></td>" & vbcrlf & _
		  "<td>    " & vbcrlf & _
		    "<table cellpadding=0 cellspacing=0 border=0 width=580 align=center>" & vbcrlf & _
			  "<tr>" & vbcrlf & _
			  	"<td valign=top align=center>" & vbcrlf & _
				  "<table width=550 border=0 cellspacing=0 cellpadding=0 align=center>" & vbcrlf & _
				  	"<tr>" & vbcrlf & _
					  "<td valign=top>"
					  
TextBody = TextBody & "<p>Thanks for signing up! Keep this email for your records.</p>"
TextBody = TextBody & "<p>Your user name is: <b>" & strEmail & "</b></p>"
TextBody = TextBody & "<p>If you lose your password, please use the <a href=http://www.chestees.com/forgot-password/>forgot my password</a> feature on <a href=http://www.chestees.com/>chestees.com</a>.</p>"
							
							TextBody = TextBody & "</td>" & vbcrlf & _
					"</tr>" & vbcrlf & _
					"<tr>" & vbcrlf & _
					  "<td height=15><img src=http://www.chestees.com/images/filler.gif width=1 height=15></td>" & vbcrlf & _
					"</tr>" & vbcrlf & _
				  "</table>" & vbcrlf & _
				"</td>" & vbcrlf & _
			  "</tr>" & vbcrlf & _
			"</table>" & vbcrlf & _
		  "</td>" & vbcrlf & _
  		  "<td width=1 bgcolor=#000000><img src=http://www.chestees.com/images/filler.gif width=1></td>" & vbcrlf & _			
		"</tr>" & vbcrlf & _
		"<tr>" & vbcrlf & _
		  "<td colspan=3><img src=http://www.chestees.com/images/email_Footer.jpg></td>" & vbcrlf & _
		"</tr>" & vbcrlf & _
	  "</table>" & vbcrlf & _
	"</body>" & vbcrlf & _
	"</html>"
	
	Dim myMail
	Set myMail = Server.CreateObject ("CDONTS.NewMail")
	myMail.From = "info@chestees.com"
	myMail.To = "info@chestees.com"
	myMail.Subject = "CHESTEES.COM ACCOUNT INFO"
	myMail.Body  = TextBody
	myMail.MailFormat = 0
	myMail.BodyFormat = 0
	myMail.Send
	set myMail=nothing
	
End Function
'_____________________________________________________________________________________________
'END CREATE ACCOUNT EMAIL

'_____________________________________________________________________________________________
'CREATE CONTACT FORM EMAIL
Function Email_Contact()
	
	TextBody = "" & vbcrlf & _
	"<html>" & vbcrlf & _
	"<head>" & vbcrlf & _
	"<title>chestees.com</title>" & vbcrlf & _
	"<style type=text/css>" & vbcrlf & _
	"<!--" & vbcrlf & _
	  "body {" & vbcrlf & _
    	"margin: 0px;" & vbcrlf & _
		"padding: 0px;" & vbcrlf & _
		"background-image:url(http://www.chestees.com/images/bgIMG.jpg);" & vbcrlf & _
    	"background-color: #2f1111;" & vbcrlf & _
	  "}" & vbcrlf & _
	  "td {" & vbcrlf & _
	  	"font-family: Georgia, Tahoma, verdana;" & vbcrlf & _
		"font-size: 14px;" & vbcrlf & _
		"color: #000000;" & vbcrlf & _
	  "}" & vbcrlf & _
	"-->" & vbcrlf & _
	"</style>" & vbcrlf & _
	"</head>" & vbcrlf & _
	"<body leftmargin=0 topmargin=0 marginWidth=0 marginHeight=0>" & vbcrlf & _
	  "<table width=582 border=0 cellspacing=0 cellpadding=0 align=center bgcolor=#FFFFFF>" & vbcrlf & _
  		"<tr>" & vbcrlf & _
  		  "<td width=1 bgcolor=#000000><img src=http://www.chestees.com/images/filler.gif width=1></td>" & vbcrlf & _
 	 	  "<td width=580><img src=http://www.chestees.com/images/header_Contact.jpg></td>" & vbcrlf & _
		  "<td width=1 bgcolor=#000000><img src=http://www.chestees.com/images/filler.gif width=1></td>" & vbcrlf & _
  		"</tr>" & vbcrlf & _
	  	"<tr>" & vbcrlf & _
  		  "<td width=1 bgcolor=#000000><img src=http://www.chestees.com/images/filler.gif width=1></td>" & vbcrlf & _
		  "<td>    " & vbcrlf & _
		    "<table cellpadding=0 cellspacing=0 border=0 width=580 align=center>" & vbcrlf & _
			  "<tr>" & vbcrlf & _
			  	"<td valign=top align=center>" & vbcrlf & _
				  "<table width=550 border=0 cellspacing=0 cellpadding=0 align=center>" & vbcrlf & _
				  	"<tr>" & vbcrlf & _
					  "<td valign=top>"
					  
TextBody = TextBody & "<p>Email: <b>" & strEmail & "</b></p>"
TextBody = TextBody & "<p>Comment/Question: <b>" & Replace(strComment,"''","'") & "</b></p>"
							
							TextBody = TextBody & "</td>" & vbcrlf & _
					"</tr>" & vbcrlf & _
					"<tr>" & vbcrlf & _
					  "<td height=15><img src=http://www.chestees.com/images/filler.gif width=1 height=15></td>" & vbcrlf & _
					"</tr>" & vbcrlf & _
				  "</table>" & vbcrlf & _
				"</td>" & vbcrlf & _
			  "</tr>" & vbcrlf & _
			"</table>" & vbcrlf & _
		  "</td>" & vbcrlf & _
  		  "<td width=1 bgcolor=#000000><img src=http://www.chestees.com/images/filler.gif width=1></td>" & vbcrlf & _			
		"</tr>" & vbcrlf & _
	  "</table>" & vbcrlf & _
	"</body>" & vbcrlf & _
	"</html>"

	Dim myMail
	Set myMail = Server.CreateObject ("CDONTS.NewMail")
	myMail.From = "info@chestees.com"
	myMail.To = "info@chestees.com"
	myMail.Subject = "CHESTEES.COM CONTACT FORM"
	myMail.Body  = TextBody
	myMail.MailFormat = 0
	myMail.BodyFormat = 0
	myMail.Send
	set myMail=nothing
	
End Function
'_____________________________________________________________________________________________
'END CONTACT FORM EMAIL

'_____________________________________________________________________________________________
'CREATE CONFIRM EMAIL
Function Email_Confirm()
	OpenDB()
	TextBody = "<table cellpadding='5' cellspacing='0' border='0' style='font-family:Arial, Helvetica, sans-serif; font-size:12px;'>"_
&"  <tr>"_
&"    <td><b>Order Number:</b> " & intOrderID & "</td>"_
&"  </tr>"_
&"  <tr>"_
&"    <td><b>Date Ordered:</b> " & FormatDateTime(Now()) & "</td>"_
&"  </tr>"_
&"  <tr>"_
&"    <td><b>Items Ordered:</b></td>"_
&"  </tr>"_
&"  <tr>"_
&"    <td><table cellpadding='5' cellspacing='0' border='0' style='font-family:Arial, Helvetica, sans-serif; font-size:12px; background:#EAEAEA; margin-bottom:15px;'>"_
&"    	<tr>"_
&"        	<td width='70' align='center'><b>QTY</b></td>"_
&"            <td width='235'><b>Shirt</b></td>"_
&"            <td width='120'><b>Color</b></td>"_
&"            <td width='65'><b>Size</b></td>"_
&"            <td width='90' align='center'><b>Price</b></td>"_
&"        </tr>"
	
	'_____________________________________________________________________________________________
	'CREATE THE PRODUCTS RECORDSET
	SQL = "SELECT C.CartID, P.Product, C.Price, S.Size, Y.Style, C.Quantity,  R.PurchaseAmount, R.ShippingCost, R.DiscountAmount, R.TotalAmount FROM (((((tblCart C INNER JOIN tblProduct P ON C.ProductID = P.ProductID) "
		SQL = SQL & "INNER JOIN tblProductSize S ON C.ProductSizeID = S.ProductSizeID) "
		SQL = SQL & "INNER JOIN tblProductStyle Y ON C.ProductStyleID = Y.ProductStyleID)"
		SQL = SQL & "INNER JOIN relCartToOrder O ON C.CartID = O.CartID)"
		SQL = SQL & "INNER JOIN tblOrder R ON O.OrderID = R.OrderID) "
		SQL = SQL & "WHERE O.OrderID = " & intOrderID
		Set rsCart = Conn.Execute(SQL)
		
		curPurchaseAmount = rsCart("PurchaseAmount")
		curShippingCost = rsCart("ShippingCost")
		intDiscountAmount = rsCart("DiscountAmount")
		curTotalAmount = rsCart("TotalAmount")
		
	If Not rsCart.EOF Then
		curTotal = 0
		intTotalQuantity = 0
		
		Do While Not rsCart.EOF
		
			intCartID = rsCart("CartID")
			strProduct = rsCart("Product")
			strStyle = rsCart("Style")
			strSize = rsCart("Size")
			curPrice = rsCart("Price")
			intQuantity = rsCart("Quantity")
			myShippingCost = rsCart("ShippingCost")
			
			intTotalQuantity = intTotalQuantity + intQuantity
			curTotal = curTotal+curPrice*intQuantity
	
			TextBody = TextBody &"<tr>"_
			&"        	<td align='center'>" & intQuantity & "</td>"_
			&"            <td>" & strProduct & "</td>"_
			&"            <td>" & strStyle & "</td>"_
			&"            <td>" & strSize & "</td>"_
			&"            <td align='center'>" & formatCurrency(curPrice,0) & "</td>"_
			&"        </tr>"
			
		rsCart.MoveNext
		Loop
	End If
	
	rsCart.Close
	Set rsCart = nothing
	
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
		
	TextBody = TextBody & "</table></td>"_
&"  </tr>"_
&"  <tr>"_
&"    <td>"_
&"		<table cellpadding='0' cellspacing='0' border='0' style='font-family:Arial, Helvetica, sans-serif; font-size:12px;'>"_
&"		  <tr><td width='250'><b>Billing Address:</b><br />"& strB_FName & " " & strB_LName & "<br />" & strB_Address & "<br />"
		If strB_Address2 <> "" Then
			TextBody = TextBody & strB_Address2 & "<br />"
		End If
	TextBody = TextBody & strB_City & ", " & strB_State & "  " & strB_Zip & "</td>"_
&"    <td width='250'><b>Shipping Address:</b><br />" & strS_FName & " " & strS_LName & "<br />" & strS_Address & "<br />"
	If strS_Address2 <> "" Then
		TextBody = TextBody & strS_Address2 & "<br />"
	End If
	TextBody = TextBody & strS_City & ", " & strS_State & "  " & strS_Zip & "</td></tr></table></td>"_
&"  </tr>"_
&"  <tr>"_
&"    <td><b>Payment:</b><br />"_
&"    	Sub-total: " & formatCurrency(curPurchaseAmount,0) & "<br />"
If intDiscountAmount > 0 Then
		TextBody = TextBody & "Discount: " & formatCurrency(intDiscountAmount,2) & "<br />"
End If
	TextBody = TextBody & "      Shipping: " & formatCurrency(curShippingCost,2) & "<br />"_
&"      Total charged to card: " & formatCurrency(curTotalAmount,2) & "</td>"_
&"  </tr>"_
&"  <tr>"_
&"    <td align='center'>- Thank You! -<br />- Love, Chestees -</td>"_
&"  </tr>"_
&"</table>"
		
	CloseDB()
	
	strRequestXML = "<?xml version=" &chr(34) & "1.0" & chr(34)& "?><exacttarget><authorization>"_
		& "<username>"&ET_User&"</username>"_
		& "<password>"&ET_Password&"</password>"_
	  & "</authorization>"_
	  & "<system>"_
		& "<system_name>triggeredsend</system_name>"_
		& "<action>add</action>"_
		& "<TriggeredSend xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns='http://exacttarget.com/wsdl/partnerAPI'>"_
		  & "<TriggeredSendDefinition>"_
			& "<CustomerKey>Purchase</CustomerKey>"_
		  & "</TriggeredSendDefinition>"_
		  & "<Subscribers>"_
			& "<SubscriberKey>"&strEmail&"</SubscriberKey>"_
			& "<EmailAddress>"&strEmail&"</EmailAddress>"_
			& "<Attributes>"_
			  & "<Name>First Name</Name>"_
			  & "<Value>"&strB_FName&"</Value>"_
			& "</Attributes>"_
			& "<Attributes>"_
			  & "<Name>Last Name</Name>"_
			  & "<Value>"&strB_LName&"</Value>"_
			& "</Attributes>"_
			& "<Attributes>"_
			& "<Name>HTML__data</Name>"_
			& "<Value>"&Server.HTMLEncode(TextBody)&"</Value>"_
			& "</Attributes>"_
			& "<Attributes>"_
			  & "<Name>Id</Name>"_
			  & "<Value></Value>"_
			& "</Attributes>"_
		  & "</Subscribers>"_
		& "</TriggeredSend>"_
	  & "</system>"_
	& "</exacttarget>"

	Set objSXH = server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objSXH.open "POST", strAPIUrl, false

		' Post the XML body
		objSXH.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		objSXH.send ("qf=xml&xml=" & Server.URLEncode(strRequestXML))
		
		'Check the status of the request and response accordingly.
		If objSXH.status = 200 Then

		Else
			strSQLError = objXMLHttp.Status & ", " & objXMLHttp.StatusText
			Response.Write("An error occurred generating your email reciept. Please contact Chestees.<br /><br />" & strSQLError)
			OpenDB()
			SQL = "INSERT INTO tblError (Items, VisitorID, CustomerID) VALUES (" & _
				SQLEncode(strSQLError) & ", " & SQLNumEncode(cVisitorID) & ", " & SQLNumEncode(cCustomerID) & ")"
				Conn.Execute(SQL)
			CloseDB()
		End If

		Set objSXH = nothing
	
End Function
'_____________________________________________________________________________________________
'END CONFIRM EMAIL
%>
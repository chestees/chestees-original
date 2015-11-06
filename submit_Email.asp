<!--#include virtual="/includes/globalLib.asp" -->
<!--#include virtual="/includes/adovbs.inc" -->
<%
Response.Buffer = False

strEmail = Request("Email")

If inStr(strEmail,"@") > 0 AND inStr(strEmail,".") > 0 Then
	blnError = 0
Else
	blnError = 1
End If

If blnError = 0 Then

	Call OpenDB()

	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = Conn
	cmd.CommandText = "usp_IndexForm"
	cmd.CommandType = adCmdStoredProc
	
	cmd.Parameters.Append cmd.CreateParameter("Email",adVarChar,adParamInput,75)
	cmd.Parameters("Email") = strEmail
	cmd.Parameters.Append cmd.CreateParameter("EmailResult",adVarChar,adParamOutput,75)
	cmd.Parameters("EmailResult") = EmailResult
	
	cmd.Execute
	
	intEmailResult = cmd.Parameters("EmailResult")
	
	If intEmailResult = 0 Then
		
		Call CloseDB()
		
		Dim myMail
		Set myMail = Server.CreateObject ("CDONTS.NewMail")
		myMail.From = "info@chestees.com"
		myMail.To = "info@chestees.com"
		myMail.Subject = "CHESTEES.COM EMAIL SIGN-UP"
		myMail.Body  = strEmail & " has signed up on chestees.com"
		myMail.MailFormat = 1
		myMail.BodyFormat = 0
		myMail.Send
		set myMail=nothing

	End If
	
	set cmd = nothing
	
  'INSERT INTO EXACTTARGET
  strRequestXML = "<?xml version=" &chr(34) & "1.0" & chr(34)& "?><exacttarget><authorization>"_
		& "<username>"&ET_User&"</username>"_
		& "<password>"&ET_Password&"</password>"_
	  & "</authorization>"_
	  & "<system>"_
		& "<system_name>triggeredsend</system_name>"_
		& "<action>add</action>"_
		& "<TriggeredSend xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns='http://exacttarget.com/wsdl/partnerAPI'>"_
		  & "<TriggeredSendDefinition>"_
			& "<CustomerKey>EmailSignUp</CustomerKey>"_
		  & "</TriggeredSendDefinition>"_
		  & "<Subscribers>"_
			& " <SubscriberKey>"&strEmail&"</SubscriberKey>"_
			& " <EmailAddress>"&strEmail&"</EmailAddress>"_
      & " <Attributes>"_
			& "   <Name>Sign Up Form</Name>"_
			& "   <Value>1</Value>"_
			& " </Attributes>"_
			& " <Attributes>"_
			& "   <Name>Id</Name>"_
			& "   <Value></Value>"_
			& " </Attributes>"_
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
    strSQLResponse = objSXH.responsexml.xml
		OpenDB()
		SQL = "INSERT INTO tblError (Items, VisitorID, CustomerID) VALUES (" & _
			SQLEncode(strSQLResponse) & ", " & SQLNumEncode(cVisitorID) & ", " & SQLNumEncode(cCustomerID) & ")"
			Conn.Execute(SQL)
		CloseDB()
  Else
    blnError = 1
	End If

End If

If blnError = 0 Then
	response.Write("Success")
Else
	response.Write(strSQLResponse)
End If
Set objSXH = nothing
%>
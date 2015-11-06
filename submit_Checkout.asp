<!--#include virtual="/includes/globalLib.asp" -->
<!--#include virtual="/includes/adovbs.inc" -->
<%
Response.Buffer = False

'on error resume next

strErrorMessage = ""
intError = 0

'_____________________________________________________________________________________________
'REQUEST VARIABLES
	
strB_FName = trim(Request("B_FName"))
strB_LName = trim(Request("B_LName"))
strB_Email = trim(Request("B_Email"))
strB_Email_Confirm = trim(Request("B_Email_Confirm"))
strB_Address = trim(Request("B_Address"))
strB_Address2 = trim(Request("B_Address2"))
strB_City = trim(Request("B_City"))
strB_State = trim(Request("B_State"))
strB_Zip = trim(Request("B_Zip"))

blnSameAsBilling = Request("SameAsBilling")
if blnSameAsBilling = "ON" then blnSameAsBilling = 1 else blnSameAsBilling = 0

strS_FName = trim(Request("S_FName"))
strS_LName = trim(Request("S_LName"))
strS_Address = Request(trim("S_Address"))
strS_Address2 = trim(Request("S_Address2"))
strS_City = trim(Request("S_City"))
strS_State = trim(Request("S_State"))
strS_Zip = trim(Request("S_Zip"))

strCardType = Request("CardType")
strCardNumber = trim(Request("CardNumber"))
strExpMo = trim(Request("ExpMo"))
strExpYear = trim(Request("ExpYear"))
strCVV = trim(Request("CardCVV"))
strCouponCode = trim(Request("CouponCode"))

'Billing Errors
If strB_FName = "" Then
	strErrorMessage = strErrorMessage & "B_FName,"
	intError = 1
End If
If strB_LName = "" Then
	strErrorMessage = strErrorMessage & "B_LName,"
	intError = 2
End If
If strB_Email = "" Then
	strErrorMessage = strErrorMessage & "B_Email,B_Email_Confirm,"
	intError = 3
End If
If strB_Email <> strB_Email_Confirm Then
	strErrorMessage = strErrorMessage & "B_Email_Confirm,"
	intError = 4
End If
If strB_Address = "" Then
	strErrorMessage = strErrorMessage & "B_Address,"
	intError = 5
End If
If strB_City = "" Then
	strErrorMessage = strErrorMessage & "B_City,"
	intError = 6
End If
If strB_State = "0" Then
	strErrorMessage = strErrorMessage & "B_State,"
	intError = 7
End If
If strB_Zip = "" Then
	strErrorMessage = strErrorMessage & "B_Zip,"
	intError = 8
End If
'Shipping Errors
If strS_FName = "" AND blnSameAsBilling <> 1 Then
	strErrorMessage = strErrorMessage & "S_FName,"
	intError = 9
End If
If strS_LName = "" AND blnSameAsBilling <> 1 Then
	strErrorMessage = strErrorMessage & "S_LName,"
	intError = 10
End If
If strS_Address = "" AND blnSameAsBilling <> 1 Then
	strErrorMessage = strErrorMessage & "S_Address,"
	intError = 11
End If
If strS_City = "" AND blnSameAsBilling <> 1 Then
	strErrorMessage = strErrorMessage & "S_City,"
	intError = 12
End If
If strS_State = "0" AND blnSameAsBilling <> 1 Then
	strErrorMessage = strErrorMessage & "S_State,"
	intError = 13
End If
If strS_Zip = "" AND blnSameAsBilling <> 1 Then
	strErrorMessage = strErrorMessage & "S_Zip,"
	intError = 14
End If

If strS_FName = "" AND blnSameAsBilling = 0 Then
	strErrorMessage = strErrorMessage & "S_FName,"
	intError = 9
End If
If strS_LName = "" AND blnSameAsBilling = 0 Then
	strErrorMessage = strErrorMessage & "S_LName,"
	intError = 10
End If
If strS_Address = "" AND blnSameAsBilling = 0 Then
	strErrorMessage = strErrorMessage & "S_Address,"
	intError = 11
End If
If strS_City = "" AND blnSameAsBilling = 0 Then
	strErrorMessage = strErrorMessage & "S_City,"
	intError = 12
End If
If strS_State = "0" AND blnSameAsBilling = 0 Then
	strErrorMessage = strErrorMessage & "S_State,"
	intError = 13
End If
If strS_Zip = "" AND blnSameAsBilling = 0 Then
	strErrorMessage = strErrorMessage & "S_Zip,"
	intError = 14
End If

If len(strCVV) > 4 OR strCVV = "" Then
	strErrorMessage = strErrorMessage & "CardCVV,"
	intError = 15
End If
If strCardType = "0" Then
	strErrorMessage = strErrorMessage & "CardType,"
	intError = 16
End If
If strCardNumber = "" Then
	strErrorMessage = strErrorMessage & "CardNumber,"
	intError = 17
End If
If strExpMo = "0" Then
	strErrorMessage = strErrorMessage & "ExpMo,"
	intError = 18
End If
If strExpYear = "0" Then
	strErrorMessage = strErrorMessage & "ExpYear,"
	intError = 19
End If

If intError = 0 Then
	'_____________________________________________________________________________________________
	'OPEN DATABASE CONNECTION
	Call OpenDB()
	
	'CHECK IF THE USER IS ALREADY FILLED OUT THE INFO
	'BILLING TABLE
	SQL = "SELECT B." & varBuyer() & "ID FROM tblBillingAddress B INNER JOIN tblCart C ON B." & varBuyer() & "ID = C." & varBuyer() & "ID WHERE B." & varBuyer() & "ID = " & cBuyerID() & " AND C.Purchased = 0 AND B.Lock = 0"
		Set rsCheckBilling = Conn.Execute(SQL)
			
	'CHECK IF THE USER IS ALREADY FILLED OUT THE INFO
	'SHIPPING TABLE
	SQL = "SELECT S." & varBuyer() & "ID FROM tblShippingAddress S INNER JOIN tblCart C ON S." & varBuyer() & "ID = C." & varBuyer() & "ID WHERE S." & varBuyer() & "ID = " & cBuyerID() & " AND C.Purchased = 0 AND S.Lock = 0"
		Set rsCheckShipping = Conn.Execute(SQL)
	
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = Conn
	cmd.CommandText = "usp_Checkout"
	cmd.CommandType = adCmdStoredProc
	
	cmd.Parameters.Append cmd.CreateParameter("SameAsBilling",adInteger,adParamInput)
	cmd.Parameters("SameAsBilling") = blnSameAsBilling			
	
	'INSERT BILLING / INSERT SHIPPING
	If rsCheckBilling.EOF AND rsCheckShipping.EOF Then
		varOption = 1
	'UPDATE BILLING / INSERT SHIPPING
	ElseIf Not rsCheckBilling.EOF AND rsCheckShipping.EOF Then
		varOption = 2
	'INSERT BILLING / UPDATE SHIPPING
	ElseIf rsCheckBilling.EOF AND Not rsCheckShipping.EOF Then
		varOption = 3	
	'UPDATE BILLING / UPDATE SHIPPING
	ElseIf Not rsCheckBilling.EOF AND Not rsCheckShipping.EOF Then
		varOption = 4
	End If
	cmd.Parameters.Append cmd.CreateParameter("varOption",adInteger,adParamInput)
	cmd.Parameters("varOption") = varOption
	
	If cCustomerID > 0 Then
		cVisitorID = 0 
	ElseIf cVisitorID > 0 Then
		cCustomerID = 0
	End If
	
	cmd.Parameters.Append cmd.CreateParameter("VisitorID",adInteger,adParamInput)
	cmd.Parameters("VisitorID") = cVisitorID
	cmd.Parameters.Append cmd.CreateParameter("CustomerID",adInteger,adParamInput)
	cmd.Parameters("CustomerID") = cCustomerID 
	
	cmd.Parameters.Append cmd.CreateParameter("BillingFName",adVarChar,adParamInput,50)
	cmd.Parameters("BillingFName") = strB_FName
	cmd.Parameters.Append cmd.CreateParameter("BillingLName",adVarChar,adParamInput,50)
	cmd.Parameters("BillingLName") = strB_LName
	cmd.Parameters.Append cmd.CreateParameter("Email",adVarChar,adParamInput,100)
	cmd.Parameters("Email") = strB_Email
	cmd.Parameters.Append cmd.CreateParameter("BillingAddress",adVarChar,adParamInput,125)
	cmd.Parameters("BillingAddress") = strB_Address
	cmd.Parameters.Append cmd.CreateParameter("BillingAddress2",adVarChar,adParamInput,125)
	cmd.Parameters("BillingAddress2") = strB_Address2
	cmd.Parameters.Append cmd.CreateParameter("BillingCity",adVarChar,adParamInput,75)
	cmd.Parameters("BillingCity") = strB_City
	cmd.Parameters.Append cmd.CreateParameter("BillingState",adVarChar,adParamInput,2)
	cmd.Parameters("BillingState") = strB_State
	cmd.Parameters.Append cmd.CreateParameter("BillingZip",adVarChar,adParamInput,5)
	cmd.Parameters("BillingZip") = strB_Zip
	
	cmd.Parameters.Append cmd.CreateParameter("ShippingFName",adVarChar,adParamInput,50)
	cmd.Parameters("ShippingFName") = strS_FName
	cmd.Parameters.Append cmd.CreateParameter("ShippingLName",adVarChar,adParamInput,50)
	cmd.Parameters("ShippingLName") = strS_LName
	cmd.Parameters.Append cmd.CreateParameter("ShippingAddress",adVarChar,adParamInput,125)
	cmd.Parameters("ShippingAddress") = strS_Address
	cmd.Parameters.Append cmd.CreateParameter("ShippingAddress2",adVarChar,adParamInput,125)
	cmd.Parameters("ShippingAddress2") = strS_Address2
	cmd.Parameters.Append cmd.CreateParameter("ShippingCity",adVarChar,adParamInput,75)
	cmd.Parameters("ShippingCity") = strS_City
	cmd.Parameters.Append cmd.CreateParameter("ShippingState",adVarChar,adParamInput,2)
	cmd.Parameters("ShippingState") = strS_State
	cmd.Parameters.Append cmd.CreateParameter("ShippingZip",adVarChar,adParamInput,5)
	cmd.Parameters("ShippingZip") = strS_Zip
	
	'***************************************************************************************************
	
	cmd.Parameters.Append cmd.CreateParameter("CardType",adVarChar,adParamInput,50)
	cmd.Parameters("CardType") = strCardType
	cmd.Parameters.Append cmd.CreateParameter("CardNumber",adVarChar,adParamInput,50)
	cmd.Parameters("CardNumber") = strCardNumber
	cmd.Parameters.Append cmd.CreateParameter("ExpMo",adVarChar,adParamInput,2)
	cmd.Parameters("ExpMo") = strExpMo
	cmd.Parameters.Append cmd.CreateParameter("ExpYear",adVarChar,adParamInput,4)
	cmd.Parameters("ExpYear") = strExpYear
	cmd.Parameters.Append cmd.CreateParameter("CVV",adVarChar,adParamInput,4)
	cmd.Parameters("CVV") = strCVV
	cmd.Parameters.Append cmd.CreateParameter("CouponCode",adVarChar,adParamInput,50)
	cmd.Parameters("CouponCode") = Ucase(strCouponCode)
			
	cmd.Execute
	set cmd = nothing
		
	CloseDB()

End If

If intError = 0 Then
	Response.Write("Success")
Else 
	Response.Write(left(strErrorMessage,len(strErrorMessage)-1))
End If
%>
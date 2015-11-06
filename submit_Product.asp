<!--#include virtual="/includes/globalLib.asp" -->
<!--#include virtual="/includes/adovbs.inc" -->
<%
Response.Buffer = False

intError = 0

'_____________________________________________________________________________________________
'REQUEST VARIABLES
intProductID = Request("ProductID")
intProductSizeID = Request("ProductSizeID")
intProductStyleID = Request("ProductStyleID")
dtDate = Now()

'ITEM ADDED TO THE CART
If intProductSizeID > 0 AND intProductStyleID > 0 Then
	
	'_____________________________________________________________________________________________
	'OPEN DATABASE CONNECTION
	Call OpenDB()

	'_____________________________________________________________________________________________
	'GET THE PRICE OF THE PRODUCT BASED ON THE STYLE
	SQL = "SELECT Price FROM tblProductDetail WHERE ProductID = " & intProductID & _
		" AND ProductSizeID = " & intProductSizeID & _
		" AND ProductStyleID = " & intProductStyleID
		Set rsPrice = Conn.Execute(SQL)
		curPrice = rsPrice("Price")
		rsPrice.Close
		Set rsPrice = Nothing
	
	'_____________________________________________________________________________________________
	'CHECK IF THE PRODUCT IS ALREADY IN THE CART
	If cBuyerID() > 0 Then
	
		SQL = "SELECT ProductID, Quantity FROM tblCart WHERE " & varBuyer() & "ID = " & cBuyerID() & " AND ProductSizeID = " & intProductSizeID & " AND ProductStyleID = " & intProductStyleID & " AND Purchased = 0"
			Set rsCartCheck = Conn.Execute(SQL)
		
		'_____________________________________________________________________________________________
		'PRODUCT IS NOT IN THE USER'S CART SO ADD IT.
		If rsCartCheck.EOF Then
				
			SQL = "INSERT INTO tblCart (" & _
				varBuyer() & "ID, ProductID, ProductStyleID, ProductSizeID, Price, DateAdded, Quantity " & _
				") VALUES (" & _
				SQLNumEncode(cBuyerID()) & ", " & _
				SQLNumEncode(intProductID) & ", " & _
				SQLNumEncode(intProductStyleID) & ", " & _
				SQLNumEncode(intProductSizeID) & ", " & _
				SQLNumEncode(curPrice) & ", " & _
				SQLDateEncode(dtDate) & ", " & _
				"1)"
				Conn.Execute(SQL)
		
		'_____________________________________________________________________________________________
		'ADDS ONE MORE TO THE ALREADY IN THE CART ITEM
		Else
		
			intQuantity = rsCartCheck("Quantity") + 1
			rsCartCheck.close
			Set rsCartCheck = nothing
			
			SQL = "UPDATE tblCart SET Quantity = " & SQLNumEncode(intQuantity) & " WHERE " & _
				varBuyer() & "ID = " & cBuyerID() & _
				" AND ProductID = " & intProductID & _
				" AND ProductStyleID = " & intProductStyleID & _
				" AND ProductSizeID = " & intProductSizeID
				Conn.Execute(SQL)
				
		End If
	
	'_____________________________________________________________________________________________
	'CREATES AN ANONYMOUS USER AS A "VISITOR"
	Else

		SQL = "INSERT INTO tblCart (" & _
			"VisitorID, ProductID, ProductStyleID, ProductSizeID, Price, DateAdded, Quantity " & _
			") VALUES (" & _
			"-1, " & _
			SQLNumEncode(intProductID) & ", " & _
			SQLNumEncode(intProductStyleID) & ", " & _
			SQLNumEncode(intProductSizeID) & ", " & _
			SQLNumEncode(curPrice) & ", " & _
			SQLDateEncode(dtDate) & ", " & _
			"1)"
			Conn.Execute(SQL)
				
		'Gets the ID of the record just added
		SQL = "SELECT Max(CartID) AS MaxID FROM tblCart"
			Set rsMaxID = Conn.Execute(SQL)
			intVisitorID = rsMaxID("MaxID")
		
		rsMaxID.Close
		Set rsMaxID = Nothing
		
		Response.Cookies(cSiteName)("VisitorID") = cInt(intVisitorID)
		Response.Cookies(cSiteName).Expires = Now() + 1
		
		SQL = "UPDATE tblCart SET VisitorID = " & SQLNumEncode(intVisitorID) & " WHERE CartID = " & intVisitorID
			Conn.Execute(SQL)
			
	End If
	
	Call CloseDB()

ElseIf intProductSizeID = 0 OR intProductStyleID = 0 Then
	intError = 1
End If

If intError = 0 Then
	response.Write("Success")
ElseIf intError = 1 Then
	response.Write("Error")
End If
%>
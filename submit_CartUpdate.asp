<!--#include virtual="/includes/globalLib.asp" -->
<!--#include virtual="/includes/adovbs.inc" -->
<%
Response.Buffer = False

on error resume next

'_____________________________________________________________________________________________
'OPEN DATABASE CONNECTION
Call OpenDB()

If Request("Update") Then
	
	SQL = "SELECT CartID FROM tblCart WHERE " & varBuyer() & "ID = " & cBuyerID() & " AND Purchased = 0"
		Set rsQTY = Conn.Execute(SQL)
	
	If Not rsQTY.EOF Then
		Do While Not rsQTY.EOF
		
			intCartID = rsQTY("CartID")
			
			'_____________________________________________________________________________________________
			'REQUEST VARIABLES
			intQuantity = Request("Cart_" & intCartID)
			'If intQuantity = "" Then
			'	intQuantity = 0
			'End If
			
			If intQuantity > 0 Then
				
				Set cmd = Server.CreateObject("ADODB.Command")
				Set cmd.ActiveConnection = Conn
				cmd.CommandText = "usp_UpdateCart"
				cmd.CommandType = adCmdStoredProc
				
				cmd.Parameters.Append cmd.CreateParameter("CartID",adInteger,adParamInput)
				cmd.Parameters("CartID") = intCartID
				cmd.Parameters.Append cmd.CreateParameter("Quantity",adInteger,adParamInput)
				cmd.Parameters("Quantity") = intQuantity
			
				cmd.Execute
				set cmd = nothing
	
			'Else
			
			'	Set cmd = Server.CreateObject("ADODB.Command")
			'	Set cmd.ActiveConnection = Conn
			'	cmd.CommandText = "usp_DeleteCart"
			'	cmd.CommandType = adCmdStoredProc
				
			'	cmd.Parameters.Append cmd.CreateParameter("CartID",adInteger,adParamInput)
			'	cmd.Parameters("CartID") = intCartID
			
			'	cmd.Execute
			'	set cmd = nothing
			
			End If			
			
		rsQTY.MoveNext
		Loop
	End If

ElseIf Request("Delete") Then

	SQL = "DELETE FROM tblCart WHERE CartID = " & Request("CartID")
		Conn.Execute(SQL)
		 
End If
	
'_____________________________________________________________________________________________
'CREATE THE PRODUCTS SIZES RECORDSET
SQL = "SELECT C.CartID, P.Product, C.Price, S.SizeAbbr, Y.Style, C.Quantity FROM (((tblCart C INNER JOIN tblProduct P ON C.ProductID = P.ProductID) "
	SQL = SQL & "INNER JOIN tblProductSize S ON C.ProductSizeID = S.ProductSizeID) "
	SQL = SQL & "INNER JOIN tblProductStyle Y ON C.ProductStyleID = Y.ProductStyleID)"
	SQL = SQL & "WHERE C." & varBuyer() & "ID = " & cBuyerID()
	SQL = SQL & " AND C.Purchased = 0"
	Set rsCart = Conn.Execute(SQL)

If Not rsCart.EOF Then
	
	curTotal = 0
	curTotalShipping = 0
	intTotalQuantity = 0
	myColor = "#E2E2E2"
	strCart = ""
	
	Do While Not rsCart.EOF
	
		intCartID = rsCart("CartID")
		strProduct = rsCart("Product")
		strStyle = rsCart("Style")
		strSizeAbbr = rsCart("SizeAbbr")
		curPrice = rsCart("Price")
		intQuantity = rsCart("Quantity")
		
		intTotalQuantity = intTotalQuantity + intQuantity
		curTotal = curTotal+curPrice*intQuantity

		If myColor = "#E2E2E2" Then myColor = "#FFFFFF" Else myColor = "#E2E2E2"
		
		strCart = strCart & "<div id='Cart_Mod_" & intCartID & "' class='Cart_Header_Mod' style='background:" & myColor & ";'>"
		strCart = strCart & "<div class='Cart_Col1'><b>" & strProduct & "</b><br /><span class='SmallText'>Color: " & strStyle & "</span></div>"
		strCart = strCart & "<div class='Cart_Col2-3'>" & strSizeAbbr & "</div>"
		strCart = strCart & "<div class='Cart_Col2-3'>" & formatCurrency(curPrice,0) & "</div>"
		strCart = strCart & "<div class='Cart_Col4'><input id='Cart_Input_" & intCartID & "' style='text-align:center;' name='Cart_" & intCartID & "' class='Textbox_Qty' type='text' size='2' value='" & intQuantity & "'></div>"
		strCart = strCart & "<div class='Cart_Col5'><a class='small gray button delete' id='" & intCartID & "'>remove</a></div>"
		strCart = strCart & "<div class='clear'></div>"
		strCart = strCart & "</div>"
		
	rsCart.MoveNext
	Loop
	
	strCart = strCart & "<div class='Cart_Totals'>Total: " & formatCurrency(curTotal) & "</div>"

	If intTotalQuantity >= cFreeShippingNum Then
		cShippingCost = 0
	End If
	
Else
	strCart = "<div style='text-align:center; margin:10px auto 10px auto;'>"
		strCart = strCart & "Cart is empty. C'mon, fill it up!<br /><br />- <a href='/'>SHOP THE TEES</a> -"
	strCart = strCart & "</div>"
End If

Call CloseDB()

strResponse = strCart

response.Write(strResponse)
%>
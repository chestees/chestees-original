<%
Call OpenDB()

'_____________________________________________________________________________________________
'CREATE THE PRODUCTS RECORDSET
SQL = "SELECT ProductID, Product, Image_Index, LinkDesc FROM tblProduct WHERE Active = 1 AND Private = 0 AND ProductID <> 21"
	Set rsProducts=Server.CreateObject("ADODB.Recordset") 
	rsProducts.Open SQL,Conn,3,3
	
	intRecordCount = rsProducts.recordcount
	
	Randomize()
	intRandom = int(intRecordCount * Rnd())
%>
		  	  <div style="text-align:center; margin-bottom:5px; font-weight:bold;">more funny t-shirts</div>
<%	
	If intRandom = 0 Then intRandom = 1
	If intRandom = intRecordCount Then intRandom = intRandom - 1
	If Not rsProducts.EOF Then
		i = 0
		j = 0
		Do While Not rsProducts.EOF
			i = i+1
			
			intProductID = rsProducts("ProductID")
			
			If i = intRandom AND j < 3 Then
				j = j+1
				i = i-1
				strProduct = rsProducts("Product")
				strDirectory = Stripper(strProduct)
				strImage_Main = rsProducts("Image_Index")
				strLinkDesc = rsProducts("LinkDesc")
				Response.Write("<a href='/funny-t-shirts/" & intProductID & "/" & strDirectory & "/'><img class='Product_Random' src='/uploads/products/" & strImage_Main & "' alt='" & strProduct & "' title='" & strProduct & "'></a>")
				
			End If
			
		rsProducts.MoveNext
		Loop

	End If
	
	rsProducts.Close
	Set rsProducts = Nothing
	
Call CloseDB()
%>
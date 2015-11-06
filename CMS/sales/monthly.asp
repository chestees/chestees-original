<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="/includes/globalLib.asp"-->
<%
'_____________________________________________________________________________________________
'Get Variables
btnSubmit = Request("Submit")
dtYear = year(now)

If btnSubmit <> "" Then

	dtYear = Request("Year")	

End If
i = 0

Call OpenDB()
Dim myFixedArray0(11) 'Fixed size array
myFixedArray0(0) = "DateOrdered >= '" & dtYear & "-01-01' AND DateOrdered <= '" & dtYear & "-01-31 11:59:59 PM'"
myFixedArray0(1) = "DateOrdered >= '" & dtYear & "-02-01' AND DateOrdered <= '" & dtYear & "-02-28 11:59:59 PM'"
myFixedArray0(2) = "DateOrdered >= '" & dtYear & "-03-01' AND DateOrdered <= '" & dtYear & "-03-31 11:59:59 PM'"
myFixedArray0(3) = "DateOrdered >= '" & dtYear & "-04-01' AND DateOrdered <= '" & dtYear & "-04-30 11:59:59 PM'"
myFixedArray0(4) = "DateOrdered >= '" & dtYear & "-05-01' AND DateOrdered <= '" & dtYear & "-05-31 11:59:59 PM'"
myFixedArray0(5) = "DateOrdered >= '" & dtYear & "-06-01' AND DateOrdered <= '" & dtYear & "-06-30 11:59:59 PM'"
myFixedArray0(6) = "DateOrdered >= '" & dtYear & "-07-01' AND DateOrdered <= '" & dtYear & "-07-31 11:59:59 PM'"
myFixedArray0(7) = "DateOrdered >= '" & dtYear & "-08-01' AND DateOrdered <= '" & dtYear & "-08-31 11:59:59 PM'"
myFixedArray0(8) = "DateOrdered >= '" & dtYear & "-09-01' AND DateOrdered <= '" & dtYear & "-09-30 11:59:59 PM'"
myFixedArray0(9) = "DateOrdered >= '" & dtYear & "-10-01' AND DateOrdered <= '" & dtYear & "-10-31 11:59:59 PM'"
myFixedArray0(10) = "DateOrdered >= '" & dtYear & "-11-01' AND DateOrdered <= '" & dtYear & "-11-30 11:59:59 PM'"
myFixedArray0(11) = "DateOrdered >= '" & dtYear & "-12-01' AND DateOrdered <= '" & dtYear & "-12-31 11:59:59 PM'"

Dim myFixedArray1(11) 'Fixed size array
myFixedArray1(0) = "DateOrdered >= '" & dtYear-1 & "-01-01' AND DateOrdered <= '" & dtYear-1 & "-01-31 11:59:59 PM'"
myFixedArray1(1) = "DateOrdered >= '" & dtYear-1 & "-02-01' AND DateOrdered <= '" & dtYear-1 & "-02-28 11:59:59 PM'"
myFixedArray1(2) = "DateOrdered >= '" & dtYear-1 & "-03-01' AND DateOrdered <= '" & dtYear-1 & "-03-31 11:59:59 PM'"
myFixedArray1(3) = "DateOrdered >= '" & dtYear-1 & "-04-01' AND DateOrdered <= '" & dtYear-1 & "-04-30 11:59:59 PM'"
myFixedArray1(4) = "DateOrdered >= '" & dtYear-1 & "-05-01' AND DateOrdered <= '" & dtYear-1 & "-05-31 11:59:59 PM'"
myFixedArray1(5) = "DateOrdered >= '" & dtYear-1 & "-06-01' AND DateOrdered <= '" & dtYear-1 & "-06-30 11:59:59 PM'"
myFixedArray1(6) = "DateOrdered >= '" & dtYear-1 & "-07-01' AND DateOrdered <= '" & dtYear-1 & "-07-31 11:59:59 PM'"
myFixedArray1(7) = "DateOrdered >= '" & dtYear-1 & "-08-01' AND DateOrdered <= '" & dtYear-1 & "-08-31 11:59:59 PM'"
myFixedArray1(8) = "DateOrdered >= '" & dtYear-1 & "-09-01' AND DateOrdered <= '" & dtYear-1 & "-09-30 11:59:59 PM'"
myFixedArray1(9) = "DateOrdered >= '" & dtYear-1 & "-10-01' AND DateOrdered <= '" & dtYear-1 & "-10-31 11:59:59 PM'"
myFixedArray1(10) = "DateOrdered >= '" & dtYear-1 & "-11-01' AND DateOrdered <= '" & dtYear-1 & "-11-30 11:59:59 PM'"
myFixedArray1(11) = "DateOrdered >= '" & dtYear-1 & "-12-01' AND DateOrdered <= '" & dtYear-1 & "-12-31 11:59:59 PM'"

Dim myFixedArray2(11) 'Fixed size array
myFixedArray2(0) = "DateOrdered >= '" & dtYear-2 & "-01-01' AND DateOrdered <= '" & dtYear-2 & "-01-31 11:59:59 PM'"
myFixedArray2(1) = "DateOrdered >= '" & dtYear-2 & "-02-01' AND DateOrdered <= '" & dtYear-2 & "-02-28 11:59:59 PM'"
myFixedArray2(2) = "DateOrdered >= '" & dtYear-2 & "-03-01' AND DateOrdered <= '" & dtYear-2 & "-03-31 11:59:59 PM'"
myFixedArray2(3) = "DateOrdered >= '" & dtYear-2 & "-04-01' AND DateOrdered <= '" & dtYear-2 & "-04-30 11:59:59 PM'"
myFixedArray2(4) = "DateOrdered >= '" & dtYear-2 & "-05-01' AND DateOrdered <= '" & dtYear-2 & "-05-31 11:59:59 PM'"
myFixedArray2(5) = "DateOrdered >= '" & dtYear-2 & "-06-01' AND DateOrdered <= '" & dtYear-2 & "-06-30 11:59:59 PM'"
myFixedArray2(6) = "DateOrdered >= '" & dtYear-2 & "-07-01' AND DateOrdered <= '" & dtYear-2 & "-07-31 11:59:59 PM'"
myFixedArray2(7) = "DateOrdered >= '" & dtYear-2 & "-08-01' AND DateOrdered <= '" & dtYear-2 & "-08-31 11:59:59 PM'"
myFixedArray2(8) = "DateOrdered >= '" & dtYear-2 & "-09-01' AND DateOrdered <= '" & dtYear-2 & "-09-30 11:59:59 PM'"
myFixedArray2(9) = "DateOrdered >= '" & dtYear-2 & "-10-01' AND DateOrdered <= '" & dtYear-2 & "-10-31 11:59:59 PM'"
myFixedArray2(10) = "DateOrdered >= '" & dtYear-2 & "-11-01' AND DateOrdered <= '" & dtYear-2 & "-11-30 11:59:59 PM'"
myFixedArray2(11) = "DateOrdered >= '" & dtYear-2 & "-12-01' AND DateOrdered <= '" & dtYear-2 & "-12-31 11:59:59 PM'"
%>
<script type="text/javascript" src="https://www.google.com/jsapi"></script>
<script type="text/javascript">
  google.load("visualization", "1", {packages:["corechart"]});
  google.setOnLoadCallback(drawChart);
  function drawChart() {
	var data = new google.visualization.DataTable();
	data.addColumn('string', 'Month');
	data.addColumn('number', 'Purchase Amount');
	data.addColumn('number', 'Shipping');
	data.addColumn('number', 'Discount');
	data.addColumn('number', 'Total Amount');
	data.addRows(12);
<%
For Each item In myFixedArray0
	
	curPurchaseAmount_2 = 0
	cShippingCost_2 = 0
	curDiscountAmount_2 = 0
	curTotalAmount_2 = 0
	'_____________________________________________________________________________________________
	'CREATE THE PRODUCTS RECORDSET
	SQL = "SELECT PurchaseAmount, ShippingCost, DiscountAmount, TotalAmount "
		SQL = SQL & " FROM tblOrder "
		SQL = SQL & "WHERE " & item & " ORDER BY DateOrdered DESC"
		Set rsCart = Conn.Execute(SQL)
		
		Do While Not rsCart.EOF
			curPurchaseAmount = rsCart("PurchaseAmount")
			cShippingCost = rsCart("ShippingCost")
			curDiscountAmount = rsCart("DiscountAmount")
			curTotalAmount = rsCart("TotalAmount")
			
			curPurchaseAmount_2 = curPurchaseAmount_2 + curPurchaseAmount
			cShippingCost_2 = cShippingCost_2 + cShippingCost
			curDiscountAmount_2 = curDiscountAmount_2 + curDiscountAmount
			curTotalAmount_2 = curTotalAmount_2 + curTotalAmount
			
		rsCart.MoveNext
		Loop
		Select Case i
			Case 0
				thisMonth = "Jan"
			Case 1
				thisMonth = "Feb"
			Case 2
				thisMonth = "Mar"
			Case 3
				thisMonth = "Apr"
			Case 4
				thisMonth = "May"
			Case 5
				thisMonth = "June"
			Case 6
				thisMonth = "July"
			Case 7
				thisMonth = "Aug"
			Case 8
				thisMonth = "Sept"
			Case 9
				thisMonth = "Oct"
			Case 10
				thisMonth = "Nov"
			Case 11
				thisMonth = "Dec"
		End Select
%>
		data.setValue(<%=i%>, 0, '<%=thisMonth%>');
		data.setValue(<%=i%>, 1, <%=curPurchaseAmount_2%>);
		data.setValue(<%=i%>, 2, <%=cShippingCost_2%>);
		data.setValue(<%=i%>, 3, <%=curDiscountAmount_2%>);
		data.setValue(<%=i%>, 4, <%=curTotalAmount_2%>);
<%
	i = i + 1
	
	curTotal_PurchaseAmount = curTotal_PurchaseAmount + curPurchaseAmount_2
	curTotal_ShippingAmount = curTotal_ShippingAmount + cShippingCost_2
	curTotal_DiscountAmount = curTotal_DiscountAmount + curDiscountAmount_2
	curTotal_TotalAmount = curTotal_TotalAmount + curTotalAmount_2
Next
%>
 		var chart = new google.visualization.LineChart(document.getElementById('chart_div'));
        chart.draw(data, {width: 591, height: 500, backgroundColor: 'eef2fc', legend:'bottom',chartArea:{width:"500", top:"35"}, hAxis:{}});
      }



	



<%
i=0
For Each item In myFixedArray1
	
	curPurchaseAmount_2 = 0
	cShippingCost_2 = 0
	curDiscountAmount_2 = 0
	curTotalAmount_2 = 0
	'_____________________________________________________________________________________________
	'CREATE THE PRODUCTS RECORDSET
	SQL = "SELECT PurchaseAmount, ShippingCost, DiscountAmount, TotalAmount "
		SQL = SQL & " FROM tblOrder "
		SQL = SQL & "WHERE " & item & " ORDER BY DateOrdered DESC"
		Set rsCart = Conn.Execute(SQL)
		
		Do While Not rsCart.EOF
			curPurchaseAmount = rsCart("PurchaseAmount")
			cShippingCost = rsCart("ShippingCost")
			curDiscountAmount = rsCart("DiscountAmount")
			curTotalAmount = rsCart("TotalAmount")
			
			curPurchaseAmount_2 = curPurchaseAmount_2 + curPurchaseAmount
			cShippingCost_2 = cShippingCost_2 + cShippingCost
			curDiscountAmount_2 = curDiscountAmount_2 + curDiscountAmount
			curTotalAmount_2 = curTotalAmount_2 + curTotalAmount
			
		rsCart.MoveNext
		Loop
		Select Case i
			Case 0
				thisMonth = "Jan"
			Case 1
				thisMonth = "Feb"
			Case 2
				thisMonth = "Mar"
			Case 3
				thisMonth = "Apr"
			Case 4
				thisMonth = "May"
			Case 5
				thisMonth = "June"
			Case 6
				thisMonth = "July"
			Case 7
				thisMonth = "Aug"
			Case 8
				thisMonth = "Sept"
			Case 9
				thisMonth = "Oct"
			Case 10
				thisMonth = "Nov"
			Case 11
				thisMonth = "Dec"
		End Select
	i = i + 1
	
	curTotal_PurchaseAmount1 = curTotal_PurchaseAmount + curPurchaseAmount_2
	curTotal_ShippingAmount1 = curTotal_ShippingAmount + cShippingCost_2
	curTotal_DiscountAmount1 = curTotal_DiscountAmount + curDiscountAmount_2
	curTotal_TotalAmount1 = curTotal_TotalAmount + curTotalAmount_2
Next
%>
	




<%
i=0
For Each item In myFixedArray2
	
	curPurchaseAmount_2 = 0
	cShippingCost_2 = 0
	curDiscountAmount_2 = 0
	curTotalAmount_2 = 0
	'_____________________________________________________________________________________________
	'CREATE THE PRODUCTS RECORDSET
	SQL = "SELECT PurchaseAmount, ShippingCost, DiscountAmount, TotalAmount "
		SQL = SQL & " FROM tblOrder "
		SQL = SQL & "WHERE " & item & " ORDER BY DateOrdered DESC"
		Set rsCart = Conn.Execute(SQL)
		
		Do While Not rsCart.EOF
			curPurchaseAmount = rsCart("PurchaseAmount")
			cShippingCost = rsCart("ShippingCost")
			curDiscountAmount = rsCart("DiscountAmount")
			curTotalAmount = rsCart("TotalAmount")
			
			curPurchaseAmount_2 = curPurchaseAmount_2 + curPurchaseAmount
			cShippingCost_2 = cShippingCost_2 + cShippingCost
			curDiscountAmount_2 = curDiscountAmount_2 + curDiscountAmount
			curTotalAmount_2 = curTotalAmount_2 + curTotalAmount
			
		rsCart.MoveNext
		Loop
		Select Case i
			Case 0
				thisMonth = "Jan"
			Case 1
				thisMonth = "Feb"
			Case 2
				thisMonth = "Mar"
			Case 3
				thisMonth = "Apr"
			Case 4
				thisMonth = "May"
			Case 5
				thisMonth = "June"
			Case 6
				thisMonth = "July"
			Case 7
				thisMonth = "Aug"
			Case 8
				thisMonth = "Sept"
			Case 9
				thisMonth = "Oct"
			Case 10
				thisMonth = "Nov"
			Case 11
				thisMonth = "Dec"
		End Select

	i = i + 1
	
	curTotal_PurchaseAmount2 = curTotal_PurchaseAmount + curPurchaseAmount_2
	curTotal_ShippingAmount2 = curTotal_ShippingAmount + cShippingCost_2
	curTotal_DiscountAmount2 = curTotal_DiscountAmount + curDiscountAmount_2
	curTotal_TotalAmount2 = curTotal_TotalAmount + curTotalAmount_2
Next
%>
			
			
			
			
			
			//YEAR COMPARISO CHART
			google.load("visualization", "1", {packages:["corechart"]});
			google.setOnLoadCallback(drawChart2);
			function drawChart2() {
			var data = new google.visualization.DataTable();
			data.addColumn('string', 'Year');
			data.addColumn('number', 'Purchase Amount');
			data.addColumn('number', 'Shipping');
			data.addColumn('number', 'Discount');
			data.addColumn('number', 'Total Amount');
			data.addRows(4);
			
			//CURRENT YEAR
			data.setValue(0, 1, <%=curTotal_PurchaseAmount%>);
			data.setValue(0, 2, <%=curTotal_ShippingAmount%>);
			data.setValue(0, 3, <%=curTotal_DiscountAmount%>);
			data.setValue(0, 4, <%=curTotal_TotalAmount%>);

			data.setValue(1, 1, <%=curTotal_PurchaseAmount1%>);
			data.setValue(1, 2, <%=curTotal_ShippingAmount1%>);
			data.setValue(1, 3, <%=curTotal_DiscountAmount1%>);
			data.setValue(1, 4, <%=curTotal_TotalAmount1%>);

			data.setValue(2, 1, <%=curTotal_PurchaseAmount2%>);
			data.setValue(2, 2, <%=curTotal_ShippingAmount2%>);
			data.setValue(2, 3, <%=curTotal_DiscountAmount2%>);
			data.setValue(2, 4, <%=curTotal_TotalAmount2%>);

      var chart = new google.visualization.BarChart(document.getElementById('chart2_div'));
      chart.draw(data, {width: 591, height: 240, backgroundColor: 'eef2fc', legend:'bottom', chartArea:{width:"500", top:"35"}
		});
      }

</script>
<html>
<head>
  <title>
    <%=cFriendlySiteName%>
    | Administration</title>
  <link rel="stylesheet" href="/css/stylesheet.css">
  <link rel="shortcut icon" href="/images/favicon.ico" type="image/x-icon">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <table width="804" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td bgcolor="#A13846" width="2">
        <img src="/images/filler.gif" width="2" height="1">
      </td>
      <td bgcolor="#EEF2FC">
        <table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td colspan="3">
              <img src="/images/header.jpg">
            </td>
          </tr>
          <tr>
            <td width="153" valign="top">
              <table border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td align="center">
                    <table border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="2">
                          <img src="/images/filler.gif" width="2" height="38">
                        </td>
                        <td width="151" class="NavTitle">
                          Navigation
                        </td>
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
            <td width="2" bgcolor="#A13846">
              <img src="/images/filler.gif" width="2" height="1">
            </td>
            <td width="645" valign="top">
              <table border="0" cellspacing="0" cellpadding="0" width="645">
                <tr>
                  <td>
                    <table border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="2">
                          <img src="/images/filler.gif" width="2" height="38">
                        </td>
                        <td width="621" class="PageTitle" align="right">
                          SALES SUMMARY
                        </td>
                        <td width="20" bgcolor="#A13846">
                          <img src="/images/filler.gif" width="20" height="8">
                        </td>
                        <td width="2">
                          <img src="/images/filler.gif" width="2" height="1">
                        </td>
                      </tr>
                      <tr>
                        <td colspan="4">
                          <img src="/images/filler.gif" width="1" height="20">
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr>
                  <td>
                    <table border="0" cellspacing="0" cellpadding="0" style="margin-left: 23px; margin-bottom: 15px;">
                      <form action="index.asp" method="post">
                      <tr>
                        <td style="width: 85px;">
                          <select name="Year">
                            <option value="2012" <%If dtYear="2012" Then Response.Write("Selected")%>>2012</option>
                            <option value="2011" <%If dtYear="2011" Then Response.Write("Selected")%>>2011</option>
                            <option value="2010" <%If dtYear="2010" Then Response.Write("Selected")%>>2010</option>
                            <option value="2009" <%If dtYear="2009" Then Response.Write("Selected")%>>2009</option>
                            <option value="2008" <%If dtYear="2008" Then Response.Write("Selected")%>>2008</option>
                            <option value="2007" <%If dtYear="2007" Then Response.Write("Selected")%>>2007</option>
                          </select>
                        </td>
                        <td style="width: 85px;">
                          <input type="submit" name="Submit" class="Submit" value="Submit" style="width: 100px;">
                        </td>
                      </tr>
                      </form>
                    </table>
                  </td>
                  </tr>
                  <tr>
                    <td>
                      <table border="0" cellspacing="0" cellpadding="4" width="599" align="center">
                        <tr>
                          <td>
                            <div id="chart_div">
                            </div>
														<p>Year Comparison</p>
                            <div id="chart2_div">
                            </div>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
      <td bgcolor="#A13846" width="2">
        <img src="/images/filler.gif" width="2" height="1">
      </td>
    </tr>
    <tr>
      <td colspan="3" bgcolor="#A13846" height="2">
        <img src="/images/filler.gif" width="2" height="2">
      </td>
    </tr>
  </table>
  <!--#include virtual="/incFooter.asp" -->
</body>
</html>

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<title>收付款单</title>
<script language="JavaScript" src="../js/jquery.js"></script>
<script language="javascript" type="text/javascript">
function CheckValue()
{	   
	if (document.getElementById('date').value==""){
	alert('请填写制单日期！');
	return false;
	}
	if (document.getElementById('Cust').value==""){
	alert('请选择客户或供应商！');
	return false;
	}
	if (document.getElementById('mon').value==""){
	document.getElementById('mon').value=0 
	}
}
</script>
</head>

<body>
<%if Request("add")="true" then %><br>
<%
if Request("type") = "FK" then
  call CheckAuthority("strAddFK")
else
  call CheckAuthority("strAddSK")
end if
%>
<form name="cash" method="post"  onSubmit="return CheckValue()" action="../action/savecashbill.asp?add=true&type=<% =Request.QueryString("type") %>">
<div align="center" class="STYLE1">
<%
If Request("type") = "FK" Then
    Response.Write("付款单登记")
Else
    Response.Write("收款单登记")
End If
%></div>
<div align="center" class="STYLE1">
<%
today = Now()
tdate = Year(today) & "-" & Right("0" & Month(today), 2) & "-" & Right("0" & Day(today), 2)
sql = "select * from t_cash where cashcode like '"&request("type")&"-"&tdate&"%' order by cashcode desc"
Set rs = server.CreateObject("adodb.recordset")
rs.Open sql, conn, 1, 1
If rs.recordcount = 0 Then
    s_billcode = request("type")&"-"&tdate&"-0001"
Else
    s_temp = rs("cashcode")
    s_temp = Right(s_temp, 4) + 1
    s_billcode = request("type")&"-"&tdate&"-"&Right("000"&s_temp, 4)
End If
close_rs(rs)
Response.Write s_billcode
Response.Write "<INPUT type=hidden name=""billcode"" value="&s_billcode&">"

if Request("type") = "FK" then
	s_type = "采购入库"
else
	s_type = "销售出库"
end if
set rsBill = Server.CreateObject("adodb.recordset")
sql = "select adddate, custname, billcode, "
sql = sql + "yfprice - isnull((select sum(yfprice) from t_bill where [check] = 1 and planbillcode = bill.billcode), 0) - isnull((select sum(money) from t_cash where billcode = bill.billcode),0) - pay as wmon "
sql = sql + "from t_bill as bill where [check] = 1 and billtype = '" &s_type& "' and billcode = '"& Request.QueryString("billcode") &"'"
rsBill.Open sql, conn,1,1
%></div><br>
<table id="tbl" align="center"  width="600">
<tr>
	<td width="15%" align="right"><%
	If Request("type") = "FK" Then
		Response.Write("付款日期:")
	Else
		Response.Write("收款日期:")
	End If
	%></td>
	<td width="35%"><input type="text" name="date" id="date" value="<%=formatdate(date)%>"><%showdate("date")%><font color="red">*</font></td>
	<td width="15%" align="right">
	<%If Request("type") = "FK" Then
		Response.Write("采购单号:")
	Else
		Response.Write("销售单号:")
	End If
	%></td>
	<td width="35%"><input type="text" name="bill" id="bill" value="<% =Request.QueryString("billcode")%>"><a href="#" onClick="JavaScript:window.open ('selectcashbill.asp?type=<%=Request("type")%>', 'newwindow', 'height=600, width=800,top=100,left=150, toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=no')"><img border="0" src="../img/choose.gif" width="21" height="17"></a>
	</td>
</tr>
<tr>
	<td align="right">
	<%If Request("type") = "FK" Then
		Response.Write("供&nbsp;应&nbsp;商:")
	Else
		Response.Write("客&nbsp;&nbsp;&nbsp;&nbsp;户:")
	End If
	%></td>
	<td><input type="text" name="Cust" id="cust" size="30" value="<% =rsBill("custname") %>"><a href="#" onClick="JavaScript:window.open ('selectcust.asp?type=<%=Request("type")%>', 'newwindow', 'height=600, width=800,top=150,left=150, toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=yes')"><img border="0" src="../img/choose.gif" width="21" height="17"></a><font color="red">*</font>
	</td>
	<td align="right" width="15%">帐&nbsp;&nbsp;&nbsp;&nbsp;户:</td>
	<td align="left" width="35%"><select name="account" id="account">
	<%
	sql_account = "select * from t_account"
	Set rs_account = server.CreateObject("adodb.recordset")
	rs_account.open sql_account,conn,1,1
	Do While rs_account.EOF = False
	Response.Write "<option value="&rs_account("account")&">"&rs_account("account")&"</option>"
	rs_account.movenext
	loop
	%></select></td>
</tr>
<tr>
	<td align="right">支付方式:</td>
	<td><select name="type">
		<%
		sql = "select * from Dict_CashType"
		Set rs_cashtype = conn.Execute(sql)
		Do While rs_cashtype.EOF = False
		Response.Write "<option value="&rs_cashtype("name")&">"&rs_cashtype("name")&"</option>"
		rs_cashtype.movenext
		Loop
		close_rs(rs_cashtype)
		%></select></td>
	<td align="right">	<%If Request("type") = "FK" Then
		Response.Write("实付金额:")
	Else
		Response.Write("实收金额:")
	End If
	%></td>
	<td><input type="text" name="money" id="mon" value="<% =rsBill("wmon") %>"></td>
</tr>

<tr>
	<td align="right">借&nbsp;&nbsp;&nbsp;&nbsp;方:</td>
	<td><%call ShowItem("debit","debit","")%></td>
	<td align="right">贷&nbsp;&nbsp;&nbsp;&nbsp;方</td>
	<td><%call ShowItem("lender","lender","")%></td>
</tr>
<tr>
	<td align="right">经&nbsp;办&nbsp;人:</td>
	<td><%call showemployee("user","")%></td>
	<td align="right">经办部门:</td>
	<td><%call showdepart("depart","")%></td>
</tr>
<tr>
	<td align="right">备&nbsp;&nbsp;&nbsp;&nbsp;注:</td>
	<td ><input type="text" name="remark" size="40"></td>
	<td align="right">制单人:</td>
    <td ><input type="text" name="maker" size="18"  readonly=""  value="<%=request.Cookies("username")%>"></td>
</tr>
</table>
<br>
<div align="center"><input type="submit" value=" 保 存 " onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  class="button" align="middle"></div>
</form>
<%
End If
If Request("add") = "false" Then
%>
<%
if Request("type") = "FK" then
  power = Authority("strEditFK")
else
  power = Authority("strEditSK")
end if
%>
<form name="cash" method="post"  onSubmit="return CheckValue()" action="../action/savecashbill.asp?add=false&type=<% =Request.QueryString("type") %>"><br>
<div align="center"><span class="STYLE1">
	  <%
If Request("type") = "FK" Then
    Response.Write("付款单登记")
Else
    Response.Write("收款单登记")
End If

%>
	  </span></div>
<div align="center"><span class="STYLE1">
	  <%
sql = "select * from t_cash where cashcode='"&Request.QueryString("cashcode")&"'"
Set rs = server.CreateObject("adodb.recordset")
rs.Open sql, conn, 1, 3
Response.Write rs("cashcode")
Response.Write "<INPUT type=hidden name=""billcode"" value="&rs("cashcode")&">"
%></span></div><br>
<table width="700" id="tbl" align="center">
<tr>
	<td width="16%" align="right">
	<%
	If Request("type") = "FK" Then
	Response.Write("付款日期:")
	Else
	Response.Write("收款日期:")
	End If
	%></td>
	<td width="30%"><input type="text" name="date" id="date" value=<%=rs("adddate")%>>
	<%showdate("date")%><font color="red">*</font></td>
	<td width="15%" align="right">
	<%If Request("type") = "FK" Then
	Response.Write("采购单号:")
	Else
	Response.Write("销售单号:")
	End If
	%></td>
	<td width="39%"><input type="text" name="bill" id="bill" value=<%=rs("billcode")%>><a href="#SelectDate" onClick="JavaScript:window.open ('selectcashbill.asp?type=<%=Request("type")%>', 'newwindow', 'height=600, width=600,,top=150,left=150, toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=yes')"><img border="0" src="../img/choose.gif" width="21" height="17"></a></td>
</tr>
<tr>
	<td align="right">
	<%If Request("type") = "FK" Then
	Response.Write("供&nbsp;应&nbsp;商:")
	Else
	Response.Write("客&nbsp;&nbsp;&nbsp;&nbsp;户:")
	End If
	%></td>
	<td><input type="text" name="Cust" id="cust" size="20" value=<%=rs("custname")%>><a href="#" onClick="JavaScript:window.open ('selectcust.asp?type=<%=Request("type")%>', 'newwindow', 'height=600, width=800,top=150,left=150, toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=yes')"><img border="0" src="../img/choose.gif" width="21" height="17"></a><font color="red">*</font></td>
	<td align="right" width="15%">帐&nbsp;&nbsp;&nbsp;&nbsp;户:</td>
	<td align="left" width="35%"><select name="account"  id="account">
	<%
	sql_account = "select * from t_account"
	Set rs_account = server.CreateObject("adodb.recordset")
	rs_account.open sql_account,conn,1,1
	Do While rs_account.EOF = False
	if rs("account")=rs_account("account") then
	Response.Write "<option value="&rs_account("account")&" selected>"&rs_account("account")&"</option>"
	else
	Response.Write "<option value="&rs_account("account")&">"&rs_account("account")&"</option>"
	end if
	rs_account.movenext
	loop
	close_rs(rs_account)
	%></select></td>
</tr>

<tr>
	<td align="right">支付方式:</td>
	<td>
	<select name="type">
	<%
	sql = "select * from Dict_CashType"
	Set rs_cashtype = conn.Execute(sql)
	Do While rs_cashtype.EOF = False
	If rs_cashtype("name") = rs("type") Then
	Response.Write "<option value="&rs_cashtype("name")&" selected>"&rs_cashtype("name")&"</option>"
	Else
	Response.Write "<option value="&rs_cashtype("name")&">"&rs_cashtype("name")&"</option>"
	End If
	rs_cashtype.movenext
	Loop
	close_rs(rs_cashtype)
	%>
	</select></td>
	<td align="right"><%If Request("type") = "FK" Then
		Response.Write("实付金额:")
	Else
		Response.Write("实收金额:")
	End If
	%></td>
	<td><input type="text" name="money" id="mon" value=<%=rs("money")%>></td>
</tr>

<tr>
	<td align="right">借&nbsp;&nbsp;&nbsp;&nbsp;方:</td>
	<td><%call ShowItem("debit","debit",rs("debit"))%></td>
	<td align="right">贷&nbsp;&nbsp;&nbsp;&nbsp;方</td>
	<td><%call ShowItem("lender","lender",rs("lender"))%></td>
</tr>
<tr>
	<td align="right">经&nbsp;办&nbsp;人:</td>
	<td><%call showemployee("user",rs("user"))%></td>
	<td align="right">经办部门:</td>
	<td><%call showdepart("depart",rs("depart"))%></td>
</tr>
<tr>
	<td align="right">备&nbsp;&nbsp;&nbsp;&nbsp;注:</td>
	<td ><input type="text" name="remark" size="40" value=<%=rs("memo")%>></td>
	<td align="right">制单人:</td>
	<td ><input type="text" name="maker" size="18" value="<%=rs("username")%>"></td>
</tr>
</table>
<br>
<%if power = "True" then%>
<div align="center"><input type="submit" class="button" value=" 保 存 " onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  align="middle">
<%end if%>
</div>
</form>
<%
close_rs(rs)
End If
endconnection
%>
</body>
</html>

</html>

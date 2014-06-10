<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<%
if Request.QueryString("type") = "fk" then
call CheckAuthority("strBillFKCount")
end if
if Request.QueryString("type") = "sk" then
call CheckAuthority("strBillSKCount")
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<title></title>
</head>

<body  style="background:#FFFFFF">
<%
s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
s_date2 = formatdate(date)
%>
<div style="background:url(../images/main-02.gif); width:100%; height:20; padding-left:0" ><img src="../images/main-01.gif"></div>
<div style="padding-left:10px">
<br>
<%if Request.QueryString("type") = "FK" then%>
<br><div align="left"><span class="STYLE1">未付账款一览</span></div><br>
<%else%>
<br><div align="left"><span class="STYLE1">未收账款一览</span></div><br>
<%end if%>
<form id="form1" name="sample" method="post" action="cashcount_bill.asp?type=<%=request.QueryString("type")%>">
<table border="0">
<tr>
<td align="center">从日期：</td>
      <td><input type="text" name="date1" size="16"  value=<%
			If Request("date1") = "" Then
				Response.Write s_date1
			Else
				Response.Write Request("date1")
			End If%>><%showdate("date1")%></td>
      <td align="center">到日期：</td>
      <td><input type="text" name="date2" size="16"  value=<%
If Request("date2") = "" Then
    Response.Write Date()
Else
    Response.Write Request("date2")
End If

%>><%showdate("date2")%></td>
<td align="center">
单号：<input type="text" name="billcode" value="<%=request.Form("billcode")%>"></td>
<td align="center" align="left" valign="top" width="120">
<input type="submit" value=" 查 找 " onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  class="button">
</td>
</tr></table>
<hr>

<%
if Request("billcode") = "" then
s_billcode = ""
else
s_billcode = " and billcode='"&request("billcode")&"'"
end if

sql_depot = "select depotname from t_user where username='"&request.cookies("username")&"'"
Set rs_depot = conn.Execute(sql_depot)
arr = split(rs_depot("depotname"),",")
if ubound(arr) <> -1 then
  for i = lbound(arr) to ubound(arr)-1
	s_depotpower = s_depotpower & "'" & arr(i) & "',"
  next
end if
sDepotname = " and depotname in ("& s_depotpower &"'a')"

sql_cust = "select custname from t_user where username='"&request.cookies("username")&"'"
Set rs_cust = conn.Execute(sql_cust)
arr = split(rs_cust("custname"),",")
temp_cust = ""
if ubound(arr) <> -1 then
  for i = lbound(arr) to ubound(arr)-1
    temp_cust = temp_cust & "'" & arr(i) & "',"
  next
end if
sCustname = " "
%>
<%if Request.QueryString("type") = "FK" then
	sql = "select billcode,adddate,custname,depotname,username,memo,zkprice,zdprice,"
	sql = sql + "ifnull((select sum(yfprice) from t_bill where t_bill.check = 1 and planbillcode = s.billcode),0) as backmoney,"
	sql = sql + "yfprice - ifnull((select sum(yfprice) from t_bill where t_bill.check = 1 and planbillcode = s.billcode),0) as yfmoney,"
	sql = sql + "ifnull((select sum(money) from t_cash where billcode = s.billcode),0) + pay as cashmoney,"
	sql = sql + "yfprice - ifnull((select sum(yfprice) from t_bill where t_bill.check = 1 and planbillcode = s.billcode),0) - ifnull((select sum(money) from t_cash where billcode = s.billcode),0) - pay as wfmoney"
	sql = sql + " from t_bill as s where adddate between '"&s_date1&"' and '"&s_date2&"' and billtype = '采购入库'"&s_billcode&sDepotname&sCustname&" order by adddate desc,billcode desc"
	call showpage(sql,"r_fkcountbill",1)
 else
	sql = "select billcode,adddate,custname,depotname,username,memo,zkprice,zdprice,"
	sql = sql + "ifnull((select sum(yfprice) from t_bill where t_bill.check = 1 and planbillcode = s.billcode),0) as backmoney,"
	sql = sql + "yfprice - ifnull((select sum(yfprice) from t_bill where t_bill.check = 1 and planbillcode = s.billcode),0) as ysmoney,"
	sql = sql + "ifnull((select sum(money) from t_cash where billcode = s.billcode),0) + pay as cashmoney,"
	sql = sql + "yfprice - ifnull((select sum(yfprice) from t_bill where t_bill.check = 1 and planbillcode = s.billcode),0) - ifnull((select sum(money) from t_cash where billcode = s.billcode),0) - pay as wsmoney"
	sql = sql + " from t_bill as s where adddate between '"&s_date1&"' and '"&s_date2&"' and billtype = '销售出库'"&s_billcode&sDepotname&sCustname&" order by adddate desc,billcode desc"
	call showpage(sql,"r_skcountbill",1)
end if
endconnection
%>
</div>
</body>
</html>

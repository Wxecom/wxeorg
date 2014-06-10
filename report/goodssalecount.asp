<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<% call CheckAuthority("strGoodsSaleCount") %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<title>采购汇总表</title>
</head>

<body bgcolor="#FFFFFF">
<div style="background:url(../images/main-02.gif); width:100%; height:20; padding-left:0" ><img src="../images/main-01.gif"></div>
<div style="padding-left:10px">
<%
s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
s_date2 = formatdate(date)
%>
<br>
<div align="left" class="STYLE1">货品销售汇总表</div>
<table align="left"><tr><td>
<form id="form1" name="sample" method="post" action="goodssalecount.asp">
  <table width="100%" border="0" id="tbl_tittle">
    <tr>
      <td width="70" align="right" height="24">仓库名称：</td>
      <td width="150">
      <%call showdepot("depot",Request.Form("depot"))%>	  
	  </td>
      <td width="70" align="right">客&nbsp;&nbsp;&nbsp;&nbsp;户：</td>
      <td width="150"><label>
      <input type="text" name="custname" id="cust" size="16" value="<%=Request.Form("custname")%>"><a href="#SelectDate" onClick="JavaScript:window.open ('../common/selectcust.asp', 'newwindow', 'height=600, width=800,top=100,left=150, toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=no')"><img border="0" src="../img/choose.gif" width="21" height="17"></a></label></td>
      <td rowspan="3"><input type="submit" name="Submit" value=" 查 询 " onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  class="button" /></td>
    </tr>
    <tr>
      <td height="26" align="right">从&nbsp;日&nbsp;期：</td>
      <td><label>
      <input type="text" name="date1" size="16" value=<%
If Request("date1") = "" Then
    Response.Write Year(Date()) & "-" & Month(Date()) & "-1"
Else
    Response.Write Request("date1")
End If

%>><%showdate("date1")%></label></td>
      <td align="right">到&nbsp;日&nbsp;期：</td>
      <td><input type="text" name="date2" size="16" value=<%
If Request("date2") = "" Then
    Response.Write Date()
Else
    Response.Write Request("date2")
End If

%>><%showdate("date2")%></td>
    </tr>
	<tr><td align="right">货品名称：</td><td><input type="text" name="goodsname" id="goodsname" size="16" value="<%=request("goodsname")%>"></td>
	<td align="right">经&nbsp;办&nbsp;人：</td><td><select name="user" id="user">
	<option value=""></option>
    <%
sql = "select * from t_Employee"
Set a = conn.Execute(sql)
Do While a.EOF = False
    s_name = a("name")
    If Request.Form("user") = a("name") Then
        Response.Write "<option value="&s_name&" selected>"&s_name&"</option>"
    Else
        Response.Write "<option value="&s_name&">"&s_name&"</option>"
    End If
    a.movenext
Loop
close_rs(a)
%></select></td></tr>
  </table>
<hr>
</form>

<%
  sql_depot = "select depotname,RDepot from t_user where username='"&request.cookies("username")&"'"
Set rs_depot = conn.Execute(sql_depot)

if rs_depot("RDepot")=false then
   sql = "select depotname from t_depot where 1=1"
   set rs=conn.execute(sql)
    Do While rs.eof=False
	  s_depotpower = s_depotpower & "'" & rs("depotname") & "',"
	  rs.movenext
     loop 
sDepotname = " and depotname in (" & s_depotpower & "'a')" 
else
arr = split(rs_depot("depotname"),",")
if ubound(arr) <> -1 then
  for i = lbound(arr) to ubound(arr)-1
	s_depotpower = s_depotpower & "'" & arr(i) & "',"
  next
end if

  sDepotname = " and depotname in (" & s_depotpower & "'a')" 

end if

If Request.Form("date1")<>"" Then
    s_date1 = Request.Form("date1")
End If
If Request.Form("date2")<>"" Then
    s_date2 = Request.Form("date2")
End If
If Request.Form("depot") = "" Then
    s_depotname = " and depotname in ("&s_depotpower&"'a')"
Else
    s_depotname = " and depotname='"&Request.Form("depot")&"'"
End If
If Request.Form("custname") = "" Then
    s_custname = ""
Else
    s_custname = " and custname='"&Request.Form("custname")&"'"
End If
If Request.Form("goodsname") = "" Then
    s_goodsname = ""
Else
    s_goodsname = " and goodsname like'%"&Request.Form("goodsname")&"%'"
End If
If Request.Form("user") = "" Then
    s_user = ""
Else
    s_user = " and username='"&Request.Form("user")&"'"
End If
'sql = "select s1.goodscode,s1.goodsname,s1.goodsunit,s1.units,ifnull(innum,0) as t_innum,ifnull(inmoney,0) as t_inmoney,ifnull(outnum,0) as t_outnum,ifnull(outmoney,0) as t_outmoney,ifnull(innum-outnum,0) as t_totalnum,ifnull(inmoney-outmoney,0) as t_totalmoney "
'sql = sql + "from (select goodscode,goodsname,goodsunit,units from s_billdetail where billtype like '销售%'"&s_goodsname&" group by goodscode,goodsname,goodsunit,units) as s1 left join "
'sql = sql + "(select goodscode,goodsname,goodsunit,units,sum(number) as innum,sum(money) as inmoney from s_billdetail where billtype='销售出库' and adddate>='"&s_date1&"' and adddate<='"&s_date2&"'"&s_depotname&s_custname&s_user&" group by goodscode,goodsname,goodsunit,units ) as s2 "
'sql = sql + "on s1.goodscode=s2.goodscode  left join "
'sql = sql + "(select goodscode,goodsname,goodsunit,units,sum(number) as outnum,sum(money) as outmoney from s_billdetail where billtype='销售退货' and AddDate>='"&s_date1&"' and AddDate<='"&s_date2&"'"&s_depotname&s_custname&s_user&" group by goodscode,goodsname,goodsunit,units) as s3 "
'sql = sql + "on s1.goodscode=s3.goodscode "

sql = "select s1.goodscode,s1.goodsname,s1.goodsunit,s1.units,ifnull(innum,0) as t_innum,ifnull(inmoney,0) as t_inmoney,ifnull(outnum,0) as t_outnum,ifnull(outmoney,0) as t_outmoney,ifnull(innum-outnum,0) as t_totalnum,ifnull(inmoney-outmoney,0) as t_totalmoney from (select goodscode,goodsname,goodsunit,units,sum(number) as innum,sum(money) as inmoney from s_billdetail where billtype='销售出库' and AddDate>='"&s_date1&"' and AddDate<='"&s_date2&"'"&s_depotname&s_goodsname&s_user&s_custname&" group by goodscode,goodsname,goodsunit,units) as s1 left join(select goodscode,goodsname,goodsunit,units,sum(number) as outnum,sum(money) as outmoney from s_billdetail where billtype='销售退货' and AddDate>='"&s_date1&"' and AddDate<='"&s_date2&"'"&s_depotname&s_goodsname&s_user&s_custname&" group by goodscode,goodsname,goodsunit,units) as s2 on s1.goodscode=s2.goodscode"

call showpage(sql,"R_GoodsSaleCount",1)

endconnection
%>
</td></tr></table>
</div>
</body>
</html>

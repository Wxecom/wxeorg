<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<% call CheckAuthority("strJXC") %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<script language="JavaScript" src="../js/jquery.js"></script>
<script language="javascript" type="text/javascript">
$(document).ready(function() {
	$("#nodename").change(function() {
		if ($("#nodename").val()==""){
		$("#typecode").attr("value","");
		}
	});
});
</script>
<title>进销存报表</title>
</head>

<body style="background:#FFFFFF">
<div style="background:url(../images/main-02.gif); width:100%; height:20; padding-left:0" ><img src="../images/main-01.gif"></div>
<div style="padding-left:10px">
<%
s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
s_date2 = formatdate(date)
%>
<br>
<div align="left" class="STYLE1">进销存报表</div>
<table align="left"><tr><td>
<form id="form1" name="sample" method="post" action="jxc.asp">
<table border="0" id="tbl_tittle">
<tr>
  <td height="26">从&nbsp;日&nbsp;期：</td>
  <td><input type="text" name="date1" size="16" value=<%If Request("date1") = "" Then
														Response.Write Year(Date()) & "-" & Month(Date()) & "-1"
														Else
														Response.Write Request("date1")
														End If%>><%showdate("date1")%></td>
  <td>到&nbsp;日&nbsp;期：</td>
  <td><input type="text" name="date2" size="16" value=<%If Request("date2") = "" Then
														Response.Write Date()
														Else
														Response.Write Request("date2")
														End If%>><%showdate("date2")%></td>
  <td width="33%" rowspan="2"><input type="submit" name="Submit"  onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  value=" 查 询 " class="button" /></td>
</tr>
<tr>
  <td>货品类别：</td>
  <td colspan="3"><input type="text" name="nodename" id="nodename" size="16" value="<%=request.Form("nodename")%>">
                  <a href="#" onClick="window.open ('../common/tree.asp?type=goods', 'newwindow', 'left=400,top=220,height=400, width=200,toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=no')"><img border="0" src="../img/choose.gif" width="21" height="17"></a>
				  <input type="hidden" name="typecode" id="typecode" size="16" value="<%=request.Form("typecode")%>"></td>
</tr>
</table>
<hr>
</form>

<%
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

sql_depot = "select depotname from t_user where username='"&request.cookies("username")&"'"
Set rs_depot = conn.Execute(sql_depot)
arr = split(rs_depot("depotname"),",")
if ubound(arr) <> -1 then
  for i = lbound(arr) to ubound(arr)-1
	s_depotpower = s_depotpower & "'" & arr(i) & "',"
  next
end if
sDepotname = " and depotname in ("& s_depotpower &"'a')"

If Request.Form("date1")<>"" Then
  s_date1 = Request.Form("date1")
End If
If Request.Form("date2")<>"" Then
  s_date2 = Request.Form("date2")
End If
if Request.Form("typecode") = "" then
  s_typecode = ""
else
  s_typecode = " and code like '"&request.Form("typecode")&"%'"
end if
sql = "select goodscode,goodsname,goodsunit,code,units,buynum,buymoney,salenum,salemoney,buynum-salenum as num from (select t_goods.GoodsCode, t_goods.GoodsName,t_goods.goodsunit,code,t_goods.units,ifnull(s1.num,0) as buynum,ifnull(s1.mon,0) as buymoney,case when s2.num is null then 0 else s2.num*-1 end as salenum,case when s2.mon is null then 0 else s2.mon*-1 end as salemoney from (t_goods LEFT JOIN (select goodscode,sum(number*flag) as num,sum(money*flag) as mon from s_billdetail where billtype like '采购%' and (AddDate)<='"&s_date2&"' and (AddDate)>='"&s_date1&"'"&sDepotname&scustname&" group by goodscode) AS s1 ON t_goods.GoodsCode = s1.goodscode) LEFT JOIN (select goodscode,sum(number*flag) as num,sum(money*flag) as mon from s_billdetail where billtype like '销售%' and (AddDate)<='"&s_date2&"' and (AddDate)>='"&s_date1&"'"&sDepotname&scustname&" group by goodscode) AS s2 ON t_goods.GoodsCode = s2.goodscode ) as vb where 1=1 "&s_typecode&""

call showpage(sql,"r_jxc",1)
endconnection
%>
</td></tr></table>
</div>
</body>
</html>

<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<% call CheckAuthority("strDepotCount") %>
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
function detail(goodscode){
openwin('depotdetail.asp?goodscode='+goodscode,900,600)
}
function depot(goodscode){
openwin('eachdepot.asp?goodscode='+goodscode,850,600)
}
</script>
<title>库存汇总</title>
</head>

<body style="background:#FFFFFF">
<div style="background:url(../images/main-02.gif); width:100%; height:20; padding-left:0" ><img src="../images/main-01.gif"></div>
<div style="padding-left:10px">
<br>
<div align="left"><span class="style1">库存汇总报表</span></div>
<table align="left"><tr><td>
<%
s_date1 = Year(Date()) & "-" & right("00"&Month(Date()),2) & "-01"
s_date2 = formatdate(date)
%>
<form id="form1" name="sample" method="post" action="depotcount.asp">
  <table border="0" id="tbl_tittle">
    <tr>
      <td width="70" height="24" align="right">仓库名称：</td>
      <td width="150">
      <%
      Response.Write "<select name=""depot"">"
	  Response.Write "<option value=""""></option>"
      sql_depot = "select depotname from t_user where username='"&request.cookies("username")&"'"
      Set rs_depot = conn.Execute(sql_depot)
      arr = split(rs_depot("depotname"),",")
      if ubound(arr) <> -1 then
        for i = lbound(arr) to ubound(arr)-1
          s_depotpower = s_depotpower & "'" & arr(i) & "',"
	    next
      end if
      sql_depot = "select * from t_depot where depotname in (" & s_depotpower & "'')"
      set rs_s_depot = conn.execute(sql_depot)
      Do While rs_s_depot.EOF = False
        if Request.Form("depot") = rs_s_depot("depotname") then
	      Response.Write "<option value="&rs_s_depot("depotname")&" selected>"&rs_s_depot("depotname")&"</option>"
	    else
	      Response.Write "<option value="&rs_s_depot("depotname")&">"&rs_s_depot("depotname")&"</option>"
	    end if
	   rs_s_depot.movenext
      loop
      Response.Write "</select>"
	  close_rs(rs_depot)
	  close_rs(rs_s_depot)
	  %>	  
	  </td>
      <td width="70" align="right">货品类别：</td>
      <td width="150"><input type="hidden" name="typecode" id="typecode" value="<%=request.Form("typecode")%>">
      <input type="text" name="nodename" id="nodename" size="16" value=<%=Request.Form("nodename")%> ><a href="#" onClick="window.open ('../common/tree.asp?type=goods', 'newwindow', 'height=400, width=200,left=600,top=200,toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=yes')"><img border="0" src="../img/choose.gif" width="21" height="17"></a></td>
      <td rowspan="3"><input type="submit" name="Submit" value=" 查 询 " onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  class="button"/></td>
    </tr>
    <tr>
      <td height="26" align="right">从&nbsp;日&nbsp;期：</td>
      <td><input type="text" name="date1" size="16"  value=<%
			If Request("date1") = "" Then
				Response.Write s_date1
			Else
				Response.Write Request("date1")
			End If%>><%showdate("date1")%></td>
      <td align="right">到&nbsp;日&nbsp;期：</td>
      <td><input type="text" name="date2" size="16"  value=<%
If Request("date2") = "" Then
    Response.Write Date()
Else
    Response.Write Request("date2")
End If

%>><%showdate("date2")%></td>
    </tr>
	<tr>
		<td align="right">货品编码：</td>
		<td><input type="text" name="goodscode" value="<%=request.Form("goodscode")%>" size="16"></td>
		<td align="right">货品名称：</td>
		<td><input type="text" name="goodsname" value="<%=request.Form("goodsname")%>" size="16"></td>
	</tr>
  </table>
  <hr>
</form>
<%
If Request.Form("date1")<>"" Then
    s_date1 = Request.Form("date1")
End If
If Request.Form("date2")<>"" Then
    s_date2 = Request.Form("date2")
End If
s_depotpower = ""
sql_depot = "select depotname from t_user where username='"&request.cookies("username")&"'"
Set rs_depot = conn.Execute(sql_depot)
arr = split(rs_depot("depotname"),",")
if ubound(arr) <> -1 then
for i = lbound(arr) to ubound(arr)-1
  s_depotpower = s_depotpower & "'" & arr(i) & "',"
next
end if

If Request.Form("depot") = "" Then
    s_depotname = " and depotname in ("&s_depotpower&"'a')"
Else
    s_depotname = " and depotname='"&Request.Form("depot")&"'"
End If
If Request.Form("typecode") = "" Then
    s_goodstype = ""
Else
    s_goodstype = " and code like '"&Request.Form("typecode")&"%'"
End If
if Request.Form("goodscode") = "" then
	s_goodscode = ""
else
	s_goodscode = " and goodscode like '%"&request.Form("goodscode")&"%'"
end if
if Request.Form("goodsname") = "" then
	s_goodsname = ""
else
	s_goodsname = " and goodsname='"&request("goodsname")&"'"
end if

sql = "select CONCAT('<a href=# onClick=detail(''',goods.goodscode,''')>显示明细</a>)&nbsp;&nbsp;CONCAT('<a href=# onClick=depot(''',goods.goodscode,''')>分仓库存</a>') as action, goods.goodscode,goodsname,goodsunit,units,endnumber,endmoney,innumber,inmoney,outnumber,outmoney,startnumber,startmoney,avgprice,(startmoney+inmoney-outmoney-endmoney) as showcolor from ("
sql = sql + "(select * from t_goods where 1=1"&s_goodstype&s_goodscode&s_goodsname&") as goods left join (select cast((case when sum(qty) is null or sum(qty) = 0 then 0 else sum(qty*price)/sum(qty) end) as float) as avgprice,goodscode from t_stock where depotname like '"& Request.Form("depot") &"%' group by goodscode) as depot on goods.goodscode = depot.goodscode "
sql = sql + "left join (select goodscode,ifnull(sum(number*flag),0) as startnumber,ifnull(sum(number*inprice*flag),0) as startmoney from s_count where s_count.check = 1 and adddate<'"&s_date1&"'"&s_depotname&" group by goodscode) as s_start on s_start.goodscode = goods.goodscode "
sql = sql + "left join (select goodscode,ifnull(sum(number),0) as outnumber,ifnull(sum(number*inprice),0) as outmoney from s_count where s_count.check = 1 and flag=-1 and adddate>='"&s_date1&"' and adddate<='"&s_date2&"'"&s_depotname&" group by goodscode) as s_out on s_out.goodscode = goods.goodscode "
sql = sql + "left join (select goodscode,ifnull(sum(number),0) as innumber,ifnull(sum(number*inprice),0) as inmoney from s_count where s_count.check = 1 and flag=1 and adddate>='"&s_date1&"' and adddate<='"&s_date2&"'"&s_depotname&" group by goodscode) as s_in on s_in.goodscode = goods.goodscode "
sql = sql + "left join (select goodscode,ifnull(sum(number*flag),0) as endnumber,ifnull(sum(number*inprice*flag),0) as endmoney from s_count where s_count.check = 1 and adddate<='"&s_date2&"'"&s_depotname&" group by goodscode) as s_end on s_end.goodscode = goods.goodscode)"

call showpage(sql,"depotcount",1)
endconnection
%>
</td></tr></table>

</div>
</body>
</html>

<!-- #include file="../inc/connExcel.asp" -->
<%
Response.AddHeader "Content-Disposition", "attachment;filename=部门领料汇总Excel.xls"
Response.ContentType  =  "application/vnd.ms-excel "
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Excel</title>
</head>
<body>
<%
sql_dpt = "select name from t_department order by id"
set rs_dpt = server.createobject("adodb.recordset")
rs_dpt.open sql_dpt,conn,1,1
s_date1=request.querystring("s_date1")
s_date2=request.querystring("s_date2")
sql=request.form("sql")
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
<table style="font-size:12px;" width="2080" id="sqd" align="center" border="1" cellpadding="2" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFEE">
<tr align="center">
	<th width="30px"></th>
	<th >货品分类</th>
  <th >货品编码</th>
  <th >货品名称</th>
  <th >规格型号</th>
  <th >单位</th>
  <th >平均单价</th>
  <th >公司神木总库</th>
	<%
	rs_dpt.movefirst
	for i=0 to rs_dpt.recordcount - 1
		response.write "<th>"&rs_dpt("name")&"</th>"
		rs_dpt.movenext
	next
	%> 
</tr>
<%
if rs.eof=false then
for i=0 to rs.recordcount -1
	If rs.bof Or rs.EOF Then Exit For
%>
<tr>
	<th><%=i+1%></th>
	<td><%=rs("goodstype")%></td>
	<td><%=rs("goodscode")%></td>
	<td><%=rs("goodsname")%></td>
	<td><%=rs("goodsunit")%></td>
	<td><%=rs("units")%></td>
	<td align="right"><%=rs("avgprice")%></td>
	<%
	sql_ult = "select ifnull(sum(number),0) as totalnum from s_billdetail where inorout='出库' and flag=-1 and adddate>='"& s_date1 &"' and adddate<='"&s_date2&"' and goodscode='"&rs("goodscode")&"' and depotname='公司神木总库'"
		set rs_ult = server.createobject("adodb.recordset")
		rs_ult.open sql_ult,conn,1,1	
		response.write "<td align=right>"&rs_ult("totalnum")&"</td>"
		rs_ult.close
	rs_dpt.movefirst
	for j=0 to rs_dpt.recordcount - 1
		sql_ult = "select ifnull(sum(number*flag),0) as totalnum from s_billdetail where billtype like '领料%' and adddate>='"& s_date1 &"' and adddate<='"&s_date2&"' and goodscode='"&rs("goodscode")&"' and custname='"&rs_dpt("name")&"'"
		set rs_ult = server.createobject("adodb.recordset")
		rs_ult.open sql_ult,conn,1,1	
		response.write "<td align=right>"&rs_ult("totalnum")*-1&"</td>"
		rs_ult.close
		rs_dpt.movenext
	next
	%>
</tr>
<%
	rs.movenext
next
end if
endconnection
%>
</table>
</body>
</html> 
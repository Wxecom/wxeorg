<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<% call CheckAuthority("strOtherDetail") %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<script type="text/javascript" src="../js/bills.js"></script>
<title>其他明细报表</title>
</head>

<body style="background:#FFFFFF">
<div style="background:url(../images/main-02.gif); width:100%; height:20; padding-left:0" ><img src="../images/main-01.gif"></div>
<div style="padding-left:10px">
<%
s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
s_date2 = formatdate(date)
%>
<br>
<div align="left" class="style1">其他明细报表</div>
<form id="form1" name="sample" method="post" action="otherdetail.asp">
  <table width="100%" border="0" id="tbl_tittle">
    <tr>
      <td width="70" height="24" align="right">仓库名称：</td>
      <td width="150">
      <%call showdepot("depot",Request.Form("depot"))%>	  
	  </td>
      <td width="80" align="right">供应商/客户：</td>
      <td width="150"><label>
      <input type="text" name="custname" id="cust" size="16" value=<%=Request.Form("custname")%> ><a href="#" onClick="JavaScript:window.open ('../common/selectcust.asp', 'newwindow', 'height=600, width=800,top=100,left=150, toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=no')"><img border="0" src="../img/choose.gif" width="21" height="17"></a></label></td>
	<td align="right" width="70">货品编码：</td><td width="150"><input type="text" name="goodscode" id="goodscode" size="16" value="<%=request.Form("goodscode")%>"></td>
      <td rowspan="3"><input type="submit" name="Submit" onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  value=" 查 询 " class="button" /></td>
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
	<td align="right">货品名称：</td><td><input type="text" name="goodsname" id="goodsname" size="16" value="<%=request.Form("goodsname")%>"></td>
    </tr>
	<td align="right">经&nbsp;办&nbsp;人：</td><td><select name="user" id="user">
	<option value=""></option>
    <%
sql = "select * from t_Employee"
Set a = conn.Execute(sql)
Do While a.EOF = False
    s_name = a("name")
	if Request("user")=s_name then
	Response.Write "<option value="&s_name&" selected>"&s_name&"</option>"
	else
    Response.Write "<option value="&s_name&">"&s_name&"</option>"
	end if
    a.movenext
Loop
close_rs(a)
%>
	</select></td>
	<td align="right">单据类型：</td><td><select name="billtype">
	                     <option value=""></option>
	                     <option value="采购入库" <%if request.Form("billtype") = "采购入库" then response.Write "selected"%>>采购入库</option>
						 <option value="采购退货" <%if request.Form("billtype") = "采购退货" then response.Write "selected"%>>采购退货</option>
						 <option value="销售出库" <%if request.Form("billtype") = "销售出库" then response.Write "selected"%>>销售出库</option>
						 <option value="销售退货" <%if request.Form("billtype") = "销售退货" then response.Write "selected"%>>销售退货</option>
						 <option value="其他入库" <%if request.Form("billtype") = "其他入库" then response.Write "selected"%>>其他入库</option>
						 <option value="其他出库" <%if request.Form("billtype") = "其他出库" then response.Write "selected"%>>其他出库</option>
						 <option value="仓库调拨" <%if request.Form("billtype") = "仓库调拨" then response.Write "selected"%>>仓库调拨</option>
						 <option value="库存盘点" <%if request.Form("billtype") = "库存盘点" then response.Write "selected"%>>库存盘点</option></select></td>
	<td align="right">货品规格：</td><td><input type="text" size="16" name="goodsunit" value="<%=request.Form("goodsunit")%>"></td>	
	</tr>
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
    s_goodsname = " and goodsname like '%"&Request.Form("goodsname")&"%'"
End If
If Request.Form("user") = "" Then
    s_user = ""
Else
    s_user = " and username='"&Request.Form("user")&"'"
End If
If Request.Form("goodsunit") = "" Then
    s_goodsunit = ""
Else
    s_goodsunit = " and goodsunit='"&request.Form("goodsunit")&"'"
End If
If Request.Form("billtype") = "" Then
    s_billtype = ""
Else
    s_billtype = " and billtype='"&request.Form("billtype")&"'"
End If
if Request.Form("goodscode") = "" then
	s_goodscode = ""
else
	s_goodscode = " and goodscode = '"& Request.Form("goodscode") &"'"
end if
sql = "select CONCAT('<a href=# onClick=EditBill(''',billcode,''')>',billcode,'</a>') as abillcode,number * flag as t_num,s_billdetail.* from s_billdetail where (AddDate)<='"&s_date2&"' and (AddDate)>='"&s_date1&"'"&s_depotname&s_custname&s_goodsname&s_user&s_goodsunit&s_billtype&s_goodscode
'response.write(sql)
call showpage(sql,"R_Detail",1)
endconnection
%>
</div>
</body>
</html>

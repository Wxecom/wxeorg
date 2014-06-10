<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<% call CheckAuthority("strSaleTop") %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<title>销售排行</title>
</head>

<body style="background:#FFFFFF;">
<div style="background:url(../images/main-02.gif); width:100%; height:20; padding-left:0" ><img src="../images/main-01.gif"></div>
<div style="padding-left:10px">
<br>
<div align="left"><span class="style1">销售排行</span></div>
<table align="left"><tr><td>
<%
s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
s_date2 = formatdate(date)
%>
<form id="form1" name="sample" method="post" action="saletop.asp">
  <table border="0" id="tbl_tittle">
    <tr>
      <td width="70" height="24" align="right">仓库名称：</td>
      <td width="150">
      <%call showdepot("depot",Request.Form("depot"))%>
	  </td>
      <td width="70" align="right">货品分类：</td>
      <td width="150"><label>
	  <input type="hidden" name="typecode" value="<%=request.Form("typecode")%>">
      <input type="text" name="nodename" id="nodename" size="16" value=<%=Request.Form("nodename")%> ><a href="#" onClick="window.open ('../common/tree.asp?type=goods', 'newwindow', 'height=400, width=200,left=600,top=200,toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=yes')"><img border="0" src="../img/choose.gif" width="21" height="17"></a></label></td>
      <td rowspan="3"><input type="submit" name="Submit" value=" 查 询 " onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  class="button" />
      </td>
    </tr>
    <tr>
      <td height="26" align="right">从&nbsp;日&nbsp;期：</td>
      <td><label>
      <input type="text" name="date1" size="16" value=<%
If Request("date1") = "" Then
    Response.Write s_date1
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
	<tr>
	  <td align="right">货品规格：</td><td><input type="text" name="goodsunit" size="16" value="<%=request.Form("goodsunit")%>"></td>
	</tr>
  </table>
<hr>
</form>

<%

s_depotpower = ""
sql_depot = "select depotname from t_user where username='"&request.cookies("username")&"'"
Set rs_depot = conn.Execute(sql_depot)
if rs_depot("RDepot")=false then
   sql = "select depotname from t_depot where 1=1"
   set rs=conn.execute(sql)
    Do While rs.eof=False
	  s_depotpower = s_depotpower & "+'" & rs("depotname") & "'+,"
	  rs.movenext
     loop 

else
arr = split(rs_depot("depotname"),",")
if ubound(arr) <> -1 then
for i = lbound(arr) to ubound(arr)-1
  s_depotpower = s_depotpower & "+'" & arr(i) & "'+,"
next
end if
end if
	  
If Request.Form("date1")<>"" Then
    s_date1 = Request.Form("date1")
End If
If Request.Form("date2")<>"" Then
    s_date2 = Request.Form("date2")
End If
If Request.Form("depot") = "" Then
    s_depotname = " and depotname in ("&s_depotpower&"+'a'+)"
Else
    s_depotname = " and depotname=+'"&Request.Form("depot")&"'+"
End If
If Request.Form("nodename") = "" Then
    s_goodstype = ""
Else
    s_goodstype = " and code like +'"&Request.Form("typecode")&"%'+"
End If
if Request.Form("goodsunit") = "" then
  s_goodsunit = ""
else
  s_goodsunit = " and goodsunit like +'%"&request.Form("goodsunit")&"%'+"
end if


'sql = "select * from (select goodscode,goodsname,"
'sql = sql + "ifnull(sum(number*flag*-1),0) as salenum,"
'sql = sql + "ifnull(sum(money*flag*-1),0) as salemon,"
'sql = sql + "ifnull(sum(number*(price-inprice)*flag*-1),0) as gain from s_billdetail where billtype like '销售%' and (AddDate)<='"&s_date2&"' and (AddDate)>='"&s_date1&"'"&s_depotname&s_goodstype&s_goodsunit&" group by goodscode,goodsname) as vb order by salenum desc"

sql = "UP_saletop @goodinfo='"&s_depotname&s_goodstype&s_goodsunit&"',@s_date1='"&s_date1&"',@s_date2='"&s_date2&"'"

'response.write(sql)
call showpage(sql,"SaleTaxis",2)
endconnection
%>
</td></tr></table>
</div>
</body>
</html>

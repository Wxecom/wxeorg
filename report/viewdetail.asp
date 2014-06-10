<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<% call CheckAuthority("strDepotDetail") %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<link rel="StyleSheet" href="../common/dtree.css" type="text/css" />
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
<title>库存明细</title>

</head>

<body style="background:#FFFFFF" leftmargin="0" topmargin="0">
<div style="background:url(../images/main-02.gif); width:100%; height:20; padding-left:0" ><img src="../images/main-01.gif"></div>
<br>

<div align="left" style="padding-left:10px"><span class="style1">库存明细报表</span></div>
<%
s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
s_date2 = formatdate(date)
%>
<form id="form1" name="sample" method="post" action="viewdetail.asp?goodscode=<%=Request.QueryString("goodscode") %>">

  <table width="100%" border="0" id="tbl_tittle">
    <tr>
      <td width="70" align="right">仓库名称：</td>
      <td width="150"><%call showdepot("depot",Request.Form("depot"))%>
      
	  </td>
      <td width="70" align="right">货品编码：</td>
      <td width="150"><input type="text" name="code"  ></td>
	  <td width="70" align="right">货品名称：</td>
      <td width="150"><input type="text" name="gname"  ></td>
      <td rowspan="3"><input type="submit" name="Submit" value=" 查 询 " class="button" /></td>
    </tr>
    <tr>
      <td height="26" width="70" align="right">从&nbsp;日&nbsp;期：</td>
      <td><input type="text" name="date1" size="16" value=<%
			If Request("date1") = "" Then
				Response.Write Year(Date()) & "-" & Month(Date()) & "-1"
			Else
				Response.Write Request("date1")
			End If%>><%showdate("date1")%></td>
      <td align="right">到&nbsp;日&nbsp;期：</td>
      <td><input type="text" name="date2" size="16" value=<%
If Request("date2") = "" Then
    Response.Write Date()
Else
    Response.Write Request("date2")
End If

%>><%showdate("date2")%></td>
<td align="right">货品规格：</td>
<td><input type="text" name="gunit"></td>
    </tr>
  </table>

  <hr>
</form>
<div style="padding-left:10px">
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
   s_depotname = " and depotname in (" & s_depotpower & "'a')"
Else
    s_depotname = " and depotname = '"&Request.Form("depot")&"'"
End If
If Request.Form("code") = "" Then
    s_goodscode = ""
Else
    s_goodscode = " and  goodscode like '%"&Request.Form("code")&"%'"
End If
If Request.Form("gname") = "" Then
    s_goodsname = ""
Else
    s_goodsname = " and  goodsname like '%"&Request.Form("gname")&"%'"
End If
If Request.Form("gunit") = "" Then
    s_goodsunit = ""
Else
    s_goodsunit = " and  goodsunit like '%"&Request.Form("gunit")&"%'"
End If



sql = "select top 1000000 adddate,billcode,billtype,goodscode,goodsname,goodsunit,Remark,DetailNote,custname,case when flag=1 then number else 0 end as inqty,case when flag=-1 then number else 0 end as outqty,price,money,maker,depotname,username from s_billdetail where s_billdetail.check=1 "&s_depotname&" "&s_goodsname&" "&s_goodscode&" "&s_goodsunit&" and adddate between '"&s_date1&"' and '"&s_date2&"'order by adddate"

call showpage(sql,"rgoodsdepot",1)

endconnection
%>
<div>
</body>
</html>

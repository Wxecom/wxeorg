<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<title></title>
<SCRIPT   language="javascript">
function del(str,check){
if(check==1){
alert('审核过的单据不允许删除！');
return false;
}
if(!confirm('确定要删除该单据吗?')){
return false
}else{
window.location.href='../action/delbill.asp?billcode='+str
}
}
function NewBill(){
openwin('outbill.asp', 900,600)
}

function EditBill(str,check){
if(check==1){
alert('审核过的单据不允许修改！');
return false;
}
openwin('editbill.asp?billcode='+str, 900,600)
}
function checkbill(str,check){
if(check==1){
alert('该单据已是审核状态，不需要审核！');
return false;
}
openwin('../action/checkbill.asp?billcode='+str,100,100)
}
function uncheckbill(str,check){
if(check==0){
alert('该单据已是未审核状态，不需要反审！');
return false;
}
openwin('../action/uncheckbill.asp?billcode='+str,100,100)
}
</SCRIPT>
<script language="JavaScript" src="../js/ShowHideDiv.js"></script>
</head>

<body style=""padding:0px;margin:0px;"" onmousemove="movediv(event)" onmouseup="obj=0">
<div style="background:url(../images/main-02.gif); width:100%; height:20; padding-left:0" ><img src="../images/main-01.gif"></div>
<div style="padding-left:10px">
<br>
<%
s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
s_date2 = formatdate(date)
%>
<p align="left"><span class="STYLE1">其他出库查询</span></p>
<table align="left"><tr><td>
<form name="form1" method="post" action="selectotherin.asp">
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
<input type="submit" value=" 查 找 " onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  class="button">
<input type="button" value="更多条件" onMouseOver="this.className='button_over';" onMouseOut="this.className='button';" class="button" onClick="show_hide_div()">
</td></tr></table>
<div id="mdiv"style="position:absolute; visibility:hidden;top:178px; left:260px;background:#E9F0F8; border:1px solid #AFB799; font-family:verdana; font-size:12px;color=#111;" >
<table width="220px" border="0" cellpadding="0" cellspacing="0" style="border:0 solid #e7e3e7;border-collapse:collapse;">
    <tr height="20" onMouseDown="finddiv(event,mdiv)" style="cursor:move;z-index:100">
        <td colspan="2" background="../img/listbar_bg.jpg" width="220px" class="bg">
            &nbsp;<font color="#FFFFFF">更多条件</font>
        </td>
	<tr>
		<td width="30%">&nbsp;单&nbsp;&nbsp;&nbsp;&nbsp;号：</td>
		<td><input type="text" name="billcode" value="<%=request("billcode")%>" size="16"></td>
	</tr>
	<tr>
		<td>&nbsp;货品名称：</td>
		<td><input type="text" name="goodscode" value="<%=request("goodscode")%>" size="16"></td>
	</tr>
	<tr>
		<td>&nbsp;货品编码：</td>
		<td><input type="text" name="goodsname" value="<%=request("goodsname")%>" size="16"></td>
	</tr>
	<tr>
		<td>&nbsp;货品规格：</td>
		<td><input type="text" name="goodsunit" value="<%=request("goodsunit")%>" size="16"></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><input type="submit" class="button" onMouseOver="this.className='button_over';" onMouseOut="this.className='button';" value="&nbsp;查&nbsp;询&nbsp;">
			   <input type="button" id="_close" class="button" value="&nbsp;关&nbsp;闭&nbsp;" onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  onClick="show_hide_div()"></td>
</tr>
</form>
</table>
</div>
<hr>
<input type="button" value="新增其他出库单" class="button" onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  onClick="NewBill();">
<%
sql_depot = "select depotname from t_user where username='"&request.cookies("username")&"'"
Set rs_depot = conn.Execute(sql_depot)
arr = split(rs_depot("depotname"),",")
if ubound(arr) <> -1 then
  for i = lbound(arr) to ubound(arr)-1
	s_depotpower = s_depotpower & "'" & arr(i) & "',"
  next
end if
sDepotname = " and depotname in ("& s_depotpower &"'')"

if Request.Form("billcode") = "" then
	sBillcode = ""
else
	sBillcode = " and billcode = '" & Request.Form("billcode") & "'"
end if
if Request.Form("goodscode") = "" then
	sGoodscode = ""
else
	sGoodscode = " and goodscode = '" & Request.Form("goodscode") & "'"
end if
if Request.Form("goodsname") = "" then
	sGoodsname = ""
else
	sGoodsname = " and goodsname like '%" & Request.Form("goodsname") & "%'"
end if
if Request.Form("goodsunit") = "" then
	sGoodsunit = ""
else
	sGoodsunit = " and goodsunit = '" & Request.Form("goodsunit") & "'"
end if
If Request.Form("date1")<>"" Then
    s_date1 = Request.Form("date1")
End If
If Request.Form("date2")<>"" Then
    s_date2 = Request.Form("date2")
End If
sql = "select distinct CONCAT('<a href=# onClick=checkbill(''',billcode+''','+convert(varchar(1),s_billdetail.check)+')>审核</a>&nbsp;&nbsp;<a href=# onClick=uncheckbill(''',billcode+''','+convert(varchar(1),s_billdetail.check)+')>反审</a>&nbsp;&nbsp;<a href=# onClick=EditBill(''',billcode+''','+convert(varchar(1),s_billdetail.check)+')>修改</a>&nbsp;&nbsp;<a href=# onClick=del(''',billcode+''','+convert(varchar(1),s_billdetail.check)+')>删除</a>&nbsp;&nbsp;<a href=detail.asp?billcode='+billcode+'>详情</a>' as action,billcode,adddate,custname,depotname,username,memo,case when s_billdetail.check=1 then '已审核' else '未审核' end as state,checkman from s_billdetail where AddDate<='"&s_date2&"' and AddDate>='"&s_date1&"' and billtype = '其他出库'"&s_cust&sBillcode&sGoodscode&sGoodsname&sGoodsunit&sDepotname&" order by adddate desc,billcode desc"

call showpage(sql,"selectother",3)
%>
<div id="search_suggest" class="billdetail" style="display: none; position:absolute;"></div>
</td></tr></table>
</div>
</body>
</html>

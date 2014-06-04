<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<%
if request.QueryString("type")="CD" then
  call CheckAuthority("strSelectCD")
else
  call CheckAuthority("strSelectXD")
end if
%>
<%
function showdate1(fieldname)
Response.Write "<input type=""image"" src=""../img/date.gif"" height=""16"" width=""18"" style=""cursor:hand"" onClick=""calendar.show("&fieldname&");return false;"">"
end function
s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
s_date2 = formatdate(date)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">

<script language="JavaScript" src="../js/ShowHideDiv.js"></script>


<script>
function EditBill(str){
openwin('editbill.asp?billcode='+str,900,600);
}
function NewBill(str){
openwin('orderbill.asp?type='+str, 900,600)
}
function choose(str1,str2,check){
if(check=='False'){
alert('该单据已是未审核状态！');
return false;
}
openwin('addbill.asp?type='+str1+'&billcode='+str2+'&chooseplan=true', 900,600)
}
function delall(){
	var iCount = 0;
	var index = 0;
	var iSucceed = 0;
	var errMsg = "";
	$("._del").each(
		function(i)
		{
			chk = $(this)
			if(chk.attr("checked") == true){
				iCount = iCount + 1;
			}
		}
	);	if(iCount == 0){
		alert("请选择要操作的单据！");
		return false;
	}
	if(confirm('确定要删除选择的'+iCount+'条单据吗?')){
		$("._del").each( 
			function(i)
			{
				chk = $(this)
				if(chk.attr("checked") == true){
					$.post("../select/delallbill.asp",{billcode:escape(chk.val())},
					function(data)
					{ 
						index = index + 1;
						if(data == "True"){
							iSucceed = iSucceed + 1;
						}else{
							errMsg = errMsg + data + "\n";
						}
						if(index == iCount){
							alert("成功删除"+iSucceed+"张单据!");
							if(iCount != iSucceed){
								alert("未成功删除单据：\n" + errMsg);
							}
							window.location.reload();
						}
					}
					);
				}
			}
		)//遍历结束
	}
}
function checkall(){
	var iCount = 0;
	var index = 0;
	var iSucceed = 0;
	var errMsg = "";
	$("._del").each(
		function(i)
		{
			chk = $(this)
			if(chk.attr("checked") == true){
				iCount = iCount + 1;
			}
		}
	);	if(iCount == 0){
		alert("请选择要操作的单据！");
		return false;
	}
	if(confirm('确定要审核共'+iCount+'张单据吗?')){
		$("._del").each( 
			function(i)
			{
				chk = $(this);
				if(chk.attr("checked") == true){
					code = chk.val();
					$.post("../select/checkbill.asp",{billcode:escape(chk.val())},
					function(data)
					{ 
						index = index + 1;
						if(data == "True"){
							iSucceed = iSucceed + 1;
						}else{
							errMsg = errMsg + data + "\n";
						}
						if(index == iCount){
							alert("成功审核"+iSucceed+"张单据!");
							if(iCount != iSucceed){
								alert("未成功单据：\n" + errMsg);
							}
							window.location.reload();
						}
					}
					);
				}
			}
		)//遍历结束
	}
}
function uncheckall(){
	var iCount = 0;
	var index = 0;
	var iSucceed = 0;
	var errMsg = "";
	$("._del").each(
		function(i)
		{
			chk = $(this)
			if(chk.attr("checked") == true){
				iCount = iCount + 1;
			}
		}
	);	if(iCount == 0){
		alert("请选择要操作的单据！");
		return false;
	}
	if(confirm('确定要反审共'+iCount+'张单据吗?')){
		$("._del").each( 
			function(i)
			{
				chk = $(this)
				if(chk.attr("checked") == true){
					$.post("../select/uncheckbill.asp",{billcode:escape(chk.val())},
					function(data)
					{
						index = index + 1;
						if(data == "True"){
							iSucceed = iSucceed + 1;
						}else{
							errMsg = errMsg + data + "\n";
						}
						if(index == iCount){
							alert("成功反审"+iSucceed+"张单据!");
							if(iCount != iSucceed){
								alert("未成功单据：\n" + errMsg);
							}
							window.location.reload();
						}
					}
					);
				}
			}
		)//遍历结束
	}
}
</script>

<title>订单浏览</title>
</head>

<body style=""padding:0px;margin:0px;  bgcolor="#FFFFFF""" onmousemove="movediv(event)" onmouseup="obj=0">
<div style="background:url(../images/main-02.gif); width:100%; height:20; padding-left:0" ><img src="../images/main-01.gif"></div>
<div style="padding-left:10px">
<br>

<diV style="margin-left:10">
<div align="left" class="STYLE1">
<%

if request.QueryString("type")="CD" then
  Response.Write("采购订单")
else
  Response.Write("销售订单")
end if
%>
</div>
<table border="0" align="left">
<tr>
<td>
<form name="form1" method="post" action="selectplan.asp?type=<%=request.QueryString("type")%>">

<div id="mdiv"style="position:absolute; visibility:hidden;top:178px;  left:260px;background:#E9F0F8; border:1px solid #AFB799; font-family:verdana; font-size:12px;color=#111;">

<table   width="220px" border="0" cellpadding="0" cellspacing="0" style="border:0 solid #e7e3e7;border-collapse:collapse;">
    <tr height="20" onMouseDown="finddiv(event,mdiv)" style="cursor:move;z-index:100">
	
        <td colspan="2" background="../img/listbar_bg.jpg" width="220px" class="bg">
            &nbsp;<font color="#FFFFFF">更多条件</font>
        </td>
	</tr>
	<tr>
		<td width="30%">&nbsp;单&nbsp;&nbsp;&nbsp;&nbsp;号：</td>
		<td><input type="text" name="billcode" value="<%=request("billcode")%>" size="16"></td>
	</tr>
	<tr>
		<td>&nbsp;制&nbsp;单&nbsp;人：</td>
		<td><input type="text" name="maker" value="<%=request("maker")%>" size="16"></td>
	</tr>
	<tr>
		<td>&nbsp;经&nbsp;办&nbsp;人：</td>
		<td><input type="text" name="user" value="<%=request("user")%>" size="16"></td>
	</tr>
	<tr>
		<td>&nbsp;审核状态：</td>
		<td><select name="check"><option value=""></option>
						<option value="1">审核</option>
						<option value="0">未审核</option></select></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><input type="submit" class="button" value="&nbsp;查&nbsp;询&nbsp;">
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			   <input type="button" id="_close" class="button" value="&nbsp;关&nbsp;闭&nbsp;" onClick="show_hide_div()"></td>

</tr>
</table>
</div>

<table align="left">
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
</tr>
<tr>
	<td align="right">仓&nbsp;&nbsp;库：</td>
	<%
	sql_dpt = "select depotname from t_depot order by id"
	set rs_dpt = server.createobject("adodb.recordset")
	rs_dpt.open sql_dpt,conn,3,3
	%>
	<td><select name="depot" id="depot">
			<option></option>
			<%rs_dpt.movefirst
			for i=0 to rs_dpt.recordcount-1 %> 
				<option value="<%=rs_dpt("depotname")%>" <%if request.form("depot")=rs_dpt("depotname") then response.write "selected" end if%>><%=rs_dpt("depotname")%></option>
		<%		rs_dpt.movenext
			next%>
		</select></td>
	<td align="right">完成情况：</td>
	<td><select name="wcqk">
				<option value=""></option>
				<option value="未完成" <%if request.form("wcqk")="未完成" then response.write "selected" end if%> >未完成</option>
				<option value="已完成" <%if request.form("wcqk")="已完成" then response.write "selected" end if%> >已完成</option>
		</select>
	</td>	
</tr>
<tr>
      <td width="50" align="right"><%if request.QueryString("type")="CD" then
	                     Response.Write("供应商：")
					   else
					     Response.Write("客 户：")
					   end if%></td>
      <td colspan="3"><input type="text" name="custname" id="cust" size="45" value=<%=Request.Form("custname")%> ><a href="#SelectDate" onClick="JavaScript:window.open ('../common/selectcust.asp', 'selectcust', 'height=600, width=800,top=100,left=150, toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=no')"><img border="0" src="../img/choose.gif" width="21" height="17"></a></td>
      <td width="80"  valign="top" >
	  <input type="submit" value=" 查 找 "  onMouseOver="this.className='button_over';" onMouseOut="this.className='button';" class="button">
	  </td>
	  <td width="80" valign="top">
	   <input type="button" value="更多条件" onMouseOver="this.className='button_over';" onMouseOut="this.className='button';" class="button" onClick="show_hide_div()">
	   </td>
</tr>
</table>
</form>


</td></td>
<tr>
<td  style="margin-left:0px" valign="top" align="left"><hr>
<input type="button" value="新增订单"  style="margin-left:0px" class="button" onClick="NewBill('<%=request.QueryString("type")%>');">
<input type="button" class="button" value="删除" onClick="delall();">
<input type="button" class="button" value="审核" onClick="checkall();">
<input type="button" class="button" value="反审" onClick="uncheckall();">
</td>
</tr>
<tr>
<td>
<%
if Request.Form("date1") <> "" then
  s_date1 = " and adddate >= '" & Request.Form("date1") & "'"
else
  s_date1 = " and adddate >= '" & s_date1 & "'"
end if
if Request.Form("date2") <> "" then
  s_date2 = " and adddate <= '" & Request.Form("date2") & "'"
else
  s_date2 = " and adddate <= '" & s_date2 & "'"
end if
if Request.Form("custname") <> "" then
  s_cust = " and custname = '" & Request.Form("custname") & "'"
else
  s_cust = ""
end if
if Request.QueryString("type")="CD" then
  s_type="采购订货"
else
  s_type="销售订货"
end if
if Request.Form("billcode") = "" then
	s_billcode = ""
else
	s_billcode = " and billcode = '" & Request.Form("billcode") & "'"
end if
if Request.Form("maker") = "" then
	s_maker = ""
else
	s_maker = " and maker = '" & Request.Form("maker") & "'"
end if
if Request.Form("user") = "" then
	s_user = ""
else
	s_user = " and username = '" & Request.Form("user") & "'"
end if
if Request.Form("check") = "" then
	s_check = ""
else
	s_check = " and t_billdetail.check = " & Request.Form("check")
end if

if Request.Form("depot") = "" then
	s_depot = ""
else
	s_depot = " and depotname ='"& Request.Form("depot") &"'"
end if

sql_depot = "select depotname,RDepot from t_user where username='"&request.cookies("username")&"'"
Set rs_depot = conn.Execute(sql_depot)
arr = split(rs_depot("depotname"),",")
if ubound(arr) <> -1 then
  for i = lbound(arr) to ubound(arr)-1
	s_depotpower = s_depotpower & "'" & arr(i) & "',"
  next
end if
if (request.cookies("username") <> "admin") and rs_depot("RDepot") then  '非admin用户没有开启往来单位资料
  sDepotname = " and depotname in (" & s_depotpower & "'a')" 
else
  sDepotname = ""
end if
if Request.Form("wcqk") = "未完成" then
	sFinished = " and ifnull((select sum(number) from t_billdetail where billcode = bill.billcode),0) > ifnull((select sum(number) from t_billdetail where billcode in (select billcode from t_bill where planbillcode = bill.billcode)),0)"
elseif Request.Form("wcqk") = "已完成" then 
	sFinished = " and ifnull((select sum(number) from t_billdetail where billcode = bill.billcode),0) <= ifnull((select sum(number) from t_billdetail where billcode in (select billcode from t_bill where planbillcode = bill.billcode)),0)"
end if

sql = "select bill.*,CONCAT('<a href=# onClick=EditBill(''',billcode,''')>',billcode,'</a>') as abillcode,"
sql = sql & "(select sum(number) from t_billdetail where billcode = bill.billcode) as orderqty,"
sql = sql & "(select sum(number) from t_billdetail where billcode in (select billcode from t_bill where planbillcode = bill.billcode)) as finishqty,"
sql = sql & "case when bill.check=1 then '<font color=#0000FF>已审核</font>' else '<font color=#FF0000>未审核</font>' end as state,checkman"
sql = sql & " from t_bill as bill where billcode like '"& Request.QueryString("type") &"%'"
sql = sql &s_date1&s_date2&s_cust&s_billcode&s_maker&s_user&s_check&s_depot&sFinished&" order by adddate desc"
call showpage(sql,"selectplan",3)
%>
</td></tr></table>
<div id="search_suggest" class="billdetail" style="display: none; position:absolute;"></div> 
</body>
</html>

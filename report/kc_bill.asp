<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<%
if request.QueryString("type")="rk" then
title="其他入库"
url="../bills/kc_depotbill.asp?type=RK"

end if
if  request.QueryString("type")="ck" then
title="其他出库"
url="../bills/kc_depotbill.asp?type=CK"
end if
if  request.QueryString("type")="db" then
url="../bills/kc_depotbill.asp?type=db&bill=bill"
title="调拨"
end if
if  request.QueryString("type")="pd" then
title="盘点"
url="../bills/kc_depotbill.asp?type=PD&bill=pbill"
end if
if  request.QueryString("type")="ll" then
title="部门领料"
url="../bills/kc_depotbill.asp?type=LL&bill=lbill"
end if
if  request.QueryString("type")="tl" then
title="部门退料"
url="../bills/kc_depotbill.asp?type=TL&bill=lbill"
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<title></title>
<SCRIPT   language="javascript">
function NewBill(url){
openwin(url, 900,600)
}

function EditBill(str){
openwin('../bills/editbill.asp?billcode='+str, 900,600)
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
</SCRIPT>
<script language="JavaScript" src="../js/ShowHideDiv.js"></script>
</head>

<body style=" padding-left:10px; background:#FFFFFF" onmousemove="movediv(event)" onmouseup="obj=0"><br>
<%
s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
s_date2 = formatdate(date)
%>
<p align="left"><span class="STYLE1"><%=title%>查询</span></p>
<table align="left"><tr><td>
<form name="form1" method="post" action="kc_bill.asp?type=<%=request.QueryString("type")%>">
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

<td align="center" valign="top" width="100">
<input type="submit" value=" 查 找 "  class="button">
</td>
<td valign="top" width="85">
<input type="button" value="更多条件" class="button" onClick="show_hide_div()">
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
		<td  valign="top"colspan="2" align="center"><input type="submit" class="button" value="&nbsp;查&nbsp;询&nbsp;">
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	 <input type="button" id="_close" class="button" value="&nbsp;关&nbsp;闭&nbsp;" onClick="show_hide_div()"></td>
</tr>
</form>
</table>
</div>
<hr>
<table>
<tr>
<td valign="top">
	<input type="button" style="margin-top:0px" value="新增<%=title%>单" class="button" onClick="NewBill('<%=url%>');">
	<input type="button" class="button" value="删除" onClick="delall();">
	<input type="button" class="button" value="审核" onClick="checkall();">
	<input type="button" class="button" value="反审" onClick="uncheckall();">
</td>
</tr>
<tr>
<td>

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

if Request.Form("billcode") = "" then
	sBillcode = ""
else
	sBillcode = " and billcode = '" & Request.Form("billcode") & "'"
end if

if Request.Form("goodscode") = "" then
	sGoodscode = ""
else
	sGoodscode = " and billcode in (select billcode from t_billdetail where goodscode = '" & Request.Form("goodscode") & "')"
end if
if Request.Form("goodsname") = "" then
	sGoodsname = ""
else
	sGoodsname = " and billcode in (select billcode from t_billdetail where goodsname like '%" & Request.Form("goodsname") & "%')"
end if
if Request.Form("goodsunit") = "" then
	sGoodsunit = ""
else
	sGoodsunit = " and billcode in (select billcode from t_billdetail where goodsunit = '" & Request.Form("goodsunit") & "')"
end if
If Request.Form("date1")<>"" Then
    s_date1 = Request.Form("date1")
End If
If Request.Form("date2")<>"" Then
    s_date2 = Request.Form("date2")
End If
if request.QueryString("type")="rk" then
sql = "select CONCAT('<a href=# onClick=EditBill(''',billcode,''')>',billcode,'</a>') as abillcode,billcode,adddate,custname,depotname,username,memo,case when t_bill.check=1 then '<font color=#0000FF>已审核</font>' else '<font color=#FF0000>未审核</font>' end as state,checkman from t_bill where AddDate<='"&s_date2&"' and AddDate>='"&s_date1&"' and billtype = '其他入库'"&s_cust&sBillcode&sGoodscode&sGoodsname&sGoodsunit&sDepotname&" order by adddate desc,billcode desc"

call showpage(sql,"selectother",3)
end if
if  request.QueryString("type")="ck" then
sql = "select CONCAT('<a href=# onClick=EditBill(''',billcode,''')>',billcode,'</a>') as abillcode,billcode,adddate,custname,depotname,username,memo,case when t_bill.check=1 then '<font color=#0000FF>已审核</font>' else '<font color=#FF0000>未审核</font>' end as state,checkman from t_bill where AddDate<='"&s_date2&"' and AddDate>='"&s_date1&"' and billtype = '其他出库'"&s_cust&sBillcode&sGoodscode&sGoodsname&sGoodsunit&sDepotname&" order by adddate desc,billcode desc"

call showpage(sql,"selectother",3)
end if
if  request.QueryString("type")="db" then
sql = "select CONCAT('<a href=# onClick=EditBill(''',billcode,''')>',billcode,'</a>') as abillcode,billcode,adddate,custname,depotname,username,memo,case when t_bill.check=1 then '<font color=#0000FF>已审核</font>' else '<font color=#FF0000>未审核</font>' end as state,checkman from t_bill where AddDate<='"&s_date2&"' and AddDate>='"&s_date1&"' and billtype = '仓库调拨'"&s_cust&sBillcode&sGoodscode&sGoodsname&sGoodsunit&sDepotname&" order by adddate desc,billcode desc"

call showpage(sql,"selectdiaobo",3)

end if
if  request.QueryString("type")="pd" then
sql = "select CONCAT('<a href=# onClick=EditBill(''',billcode,''')>',billcode,'</a>') as abillcode,billcode,adddate,custname,depotname,username,memo,checkman,case when t_bill.check=1 then '<font color=#0000FF>已审核</font>' else '<font color=#FF0000>未审核</font>' end as state from t_bill where AddDate<='"&s_date2&"' and AddDate>='"&s_date1&"' and billtype = '库存盘点'"&s_cust&sBillcode&sGoodscode&sGoodsname&sGoodsunit&sDepotname&" order by adddate desc,billcode desc"

call showpage(sql,"selectpandian",3)
end if
if  request.QueryString("type")="ll" then
sql = "select CONCAT('<a href=# onClick=EditBill(''',billcode,''')>',billcode,'</a>') as abillcode,billcode,adddate,custname,depotname,username,memo,case when t_bill.check=1 then '<font color=#0000FF>已审核</font>' else '<font color=#FF0000>未审核</font>' end as state,checkman from t_bill where AddDate<='"&s_date2&"' and AddDate>='"&s_date1&"' and billtype = '领料出库'"&s_cust&sBillcode&sGoodscode&sGoodsname&sGoodsunit&sDepotname&" order by adddate desc,billcode desc"

call showpage(sql,"selectlingliao",3)
end if
if  request.QueryString("type")="tl" then
sql = "select CONCAT('<a href=# onClick=EditBill(''',billcode,''')>',billcode,'</a>') as abillcode,billcode,adddate,custname,depotname,username,memo,case when t_bill.check=1 then '<font color=#0000FF>已审核</font>' else '<font color=#FF0000>未审核</font>' end as state,checkman from t_bill where AddDate<='"&s_date2&"' and AddDate>='"&s_date1&"' and billtype = '退料入库'"&s_cust&sBillcode&sGoodscode&sGoodsname&sGoodsunit&sDepotname&" order by adddate desc,billcode desc"

call showpage(sql,"selectlingliao",3)
end if

%>
</td></tr></table>
<div id="search_suggest" class="billdetail" style="display: none; position:absolute;"></div>

</body>
</html>

<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<% call CheckAuthority("strAllBill") %>
<html>
<head>
<script language="JavaScript" src="../js/jquery.js"></script>
<script>
function EditBill(str){
openwin('../bills/editbill.asp?billcode='+str,900,600);
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
	);
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
	);
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
	);
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
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<script language="JavaScript" src="../js/xmlHttp.js"></script>
<title>采购汇总表</title>
</head>

<body style="background:#FFFFFF">
<div style="background:url(../images/main-02.gif); width:100%; height:20; padding-left:0" ><img src="../images/main-01.gif"></div>
<div style="padding-left:10px">
<%
s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
s_date2 = formatdate(date)
%>
<br>
<div align="left"><span class="style1">单据查询</span></div>
<table align="left"><tr><td>
<form id="form1" name="sample" method="post" action="allbill.asp">
  <table width="100%" border="0" id="tbl_tittle">
    <tr>
      <td width="70" height="24" align="right">仓库名称：</td>
      <td width="150">
      <%call showdepot("depot",Request.Form("depot"))%>	  
	  </td>
      <td width="80" align="right">供应商/客户：</td>
      <td width="150"><label>
      <input type="text" name="custname" id="cust" size="16" value=<%=Request.Form("custname")%> ><a href="#" onClick="JavaScript:window.open ('../common/selectcust.asp', 'newwindow', 'height=600, width=800,top=100,left=150, toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=no')"><img border="0" src="../img/choose.gif" width="21" height="17"></a></label></td>
      <td rowspan="2" colspan="3"><input type="submit" onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  name="Submit" value=" 查 询 " class="button" /></td>
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
	<tr><td align="right">货品名称：</td><td><input type="text" name="goodsname" id="goodsname" size="16" value="<%=Request.Form("goodsname")%>"></td>
	<td align="right">经&nbsp;办&nbsp;人：</td><td><select name="user" id="user">
	                     <option value=""></option>
    <%
sql = "select * from t_employee"
Set a = conn.Execute(sql)
Do While a.EOF = False
    s_name = a("name")
	if Request("user") = s_name then
	Response.Write "<option value="&s_name&" selected>"&s_name&"</option>"
	else
    Response.Write "<option value="&s_name&">"&s_name&"</option>"
	end if
    a.movenext
Loop
close_rs(a)
%>
	</select></td>
	</tr>
	<tr>
	  <td align="right">单据号码：</td>
	  <td><input type="text" size="16" name="billcode" value="<%=request.Form("billcode")%>"></td>
	  <td align="right">单据类型：</td>
	  <td><select name="billtype">
	                     <option value=""></option>
						 <%if request.Form("billtype")="采购订货" then%>
	                     <option value="采购订货" selected="selected">采购订货</option>
						 <%else%>
						 <option value="采购订货">采购订货</option>
						 <%end if%>
						 
						 <%if request.Form("billtype")="采购入库" then%>
	                     <option value="采购入库" selected="selected">采购入库</option>
						 <%else%>
						 <option value="采购入库">采购入库</option>
						 <%end if%>
						 
						 <%if request.Form("billtype")="采购退货" then%>
						 <option value="采购退货" selected="selected">采购退货</option>
						 <%else%>
						 <option value="采购退货">采购退货</option>
						 <%end if%>
						 
						 <%if request.Form("billtype")="销售订货" then%>
	                     <option value="销售订货" selected="selected">销售订货</option>
						 <%else%>
						 <option value="销售订货">销售订货</option>
						 <%end if%>
						 
						 <%if request.Form("billtype")="销售出库" then%>
						 <option value="销售出库" selected="selected">销售出库</option>
						 <%else%>
						 <option value="销售出库">销售出库</option>
						 <%end if%>
						 
						 <%if request.Form("billtype")="销售退货" then%>
						 <option value="销售退货" selected="selected">销售退货</option>
						 <%else%>
						 <option value="销售退货">销售退货</option>
						 <%end if%>
						 
						 <%if request.Form("billtype")="其他入库" then%>
						 <option value="其他入库" selected="selected">其他入库</option>
						 <%else%>
						 <option value="其他入库">其他入库</option>
						 <%end if%>
						 
						 <%if request.Form("billtype")="其他出库" then%>
						 <option value="其他出库" selected="selected">其他出库</option>
						 <%else%>
						 <option value="其他出库">其他出库</option>
						 <%end if%>
						 
						 <%if request.Form("billtype")="仓库调拨" then%>
						 <option value="仓库调拨" selected="selected">仓库调拨</option>
						 <%else%>
						 <option value="仓库调拨">仓库调拨</option>
						 <%end if%>
						 
						 <%if request.Form("billtype")="仓库盘点" then%>
						 <option value="库存盘点" selected="selected">库存盘点</option>
						 <%else%>
						 <option value="库存盘点">库存盘点</option>
						 <%end if%>
						 </select></td>
    </tr>
	<tr><td align="right">货品编码：</td><td><input type="text" name="goodscode" id="goodscode" size="16" value="<%=Request.Form("goodscode")%>"></td>
	<td align="right">审核状态：</td><td><select name="checked" id="checked">
	                     <option value=""></option>
						 <%
						 if Request.Form("checked") = "已审核" then
						 	selected1 = "selected"
						elseif Request.Form("checked") = "未审核" then
							selected2 = "selected"
						end if
						 %>
	                     <option value="已审核" <% =selected1%>>已审核</option>
	                     <option value="未审核" <% =selected2%>>未审核</option>
	</select></td>
	</form>
  </table>
  <hr>
<table border="0">
<tr><td>
<input type="button" class="button" value="删除" onClick="delall();">
<input type="button" class="button" value="审核" onClick="checkall();">
<input type="button" class="button" value="反审" onClick="uncheckall();">
</td></tr>
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
    s_depotname = " and (depotname in ("&s_depotpower&"'a') or custname in ("&s_depotpower&"'a'))"
Else
    s_depotname = " and (depotname='"&Request.Form("depot")&"' or custname='"&Request.Form("depot")&"')"
End If
If Request.Form("custname") = "" Then
    s_custname = ""
Else
    s_custname = " and custname='"&Request.Form("custname")&"'"
End If
If Request.Form("goodsname") = "" Then
    s_goodsname = ""
Else
    s_goodsname = " and billcode in (select billcode from t_billdetail where goodsname like '%"&Request.Form("goodsname")&"%')"
End If
if Request.Form("goodscode") = "" then
	s_goodscode = ""
else
	s_goodscode = " and billcode in (select billcode from t_billdetail where goodscode = '"& Request.Form("goodscode") &"')"
end if
If Request.Form("user") = "" Then
    s_user = ""
Else
    s_user = " and username='"&Request.Form("user")&"'"
End If
If Request.Form("billcode") = "" Then
    s_billcode = ""
Else
    s_billcode = " and billcode='"&request.Form("billcode")&"'"
End If
If Request.Form("billtype") = "" Then
    s_billtype = ""
Else
    s_billtype = " and billtype='"&request.Form("billtype")&"'"
End If
if Request.Form("checked") = "" then
	s_check = ""
elseif Request.Form("checked") = "已审核" then
	s_check = " and [check] = 1"
elseif Request.Form("checked") = "未审核" then
	s_check = " and [check] = 0"
end if
sql = "select '<a href=# onClick=EditBill('''+billcode+''')>'+billcode+'</a>' as abillcode,billcode,custname,adddate,depotname,billtype,username,memo,case when [check]=1 then '<font color=#0000FF>已审核</font>' else '<font color=#FF0000>未审核</font>' end as state,checkman from t_bill where (AddDate)<='"&s_date2&"' and (AddDate)>='"&s_date1&"'"&s_depotname&s_custname&s_goodsname&s_user&s_billcode&s_billtype&s_goodscode&s_check&" order by adddate desc"
'response.write(sql)
call showpage(sql,"allbill",3)
endconnection
%>
<div id="search_suggest" class="billdetail" style="display: none; position:absolute;"></div>
</td></tr></table>
</div>
</body></html>

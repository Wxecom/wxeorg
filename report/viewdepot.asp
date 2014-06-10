<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<% call CheckAuthority("strDepotNum") %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<link rel="StyleSheet" href="../common/dtree.css" type="text/css" />
<script language="JavaScript" src="../js/ShowHideDiv.js"></script>
<script type="text/javascript" src="../js/dtree.js"></script>
<script>
function clicknode(nodename,nodecode){
	if (document.getElementById("nodename")!=null){
		document.getElementById("nodename").value = String(nodename);	
	}
	if (document.getElementById("typecode")!=null){
		document.getElementById("typecode").value = String(nodecode);	
	}	
	show_hide_div();
}
function detail(goodscode){
window.open ('depotdetail.asp?goodscode='+goodscode, 'depotdetail', 'height=600, width=900,top=100,left=150,toolbar=no,menubar=no, scrollbars=yes, resizable=no,location=no, status=yes')
}

</script>
</head>

<body style=""padding:0px;margin:0px; bgcolor="#FFFFFF" "" onmousemove="movediv(event)" onmouseup="obj=0">
<div style="background:url(../images/main-02.gif); width:100%; height:20; padding-left:0" ><img src="../images/main-01.gif"></div>
<div style="padding-left:10px">
<div id="mdiv" style="position:absolute; visibility:hidden;top:20px; left:550px;background:#E9F0F8; border:1px solid #AFB799; font-family:verdana; font-size:12px;color=#111;" >
<table width="220px" border="0" cellpadding="0" cellspacing="0" style="border:0 solid #e7e3e7;border-collapse:collapse;">
    <tr height="20" onMouseDown="finddiv(event,mdiv)" style="cursor:move;z-index:100">
        <td colspan="2" background="../img/listbar_bg.jpg" width="220px" class="bg">
            &nbsp;<font color="#FFFFFF">货品类别</font>
        </td>
    </tr>
	<tr>
    	<td>
        <div class="dtree">
        <a href="javascript: d.openAll();">全部展开</a> | <a href="javascript: d.closeAll();">全部收起</a>
	<script type='text/javascript'>
		<!--
		d = new dTree('d');
	<%		
	  Response.Write "d.add(0,-1,'货品分类','#');"
	  sqlTree="select * from t_tree where type = 'goods'"
	  set rsTree = Server.CreateObject("adodb.recordset")
	  rsTree.Open sqlTree, conn, 1, 1 
	  do while rsTree.eof=false
		i_id=rsTree("id")
		i_pid=rsTree("pid")
		s_name=rsTree("name")
		s_url=rsTree("url")
		s_code=rsTree("code")
		s_func="javascript:clicknode('"&s_name&"','"&s_code&"')"
		Response.Write "d.add("&i_id&","&i_pid&",'"&s_name&"',""javascript:clicknode('"&s_name&"','"&s_code&"')"");"
		rsTree.movenext
	  loop
	  close_rs(rsTree)
	%>
document.write(d);

//-->
</script>
	</td>
   </tr>
  </table>
</div>
<br>
<div align="left"><span class="style1">库存查询</span></div>
<table align="left"><tr><td>
<form id="form1" name="sample" method="post" action="viewdepot.asp">
  <table width="100%" border="0" id="tbl_tittle">
    <tr>
      <td width="70" align="right">仓库名称：</td>
      <td width="150"><%call showdepot("depot",Request.Form("depot"))%>
      </td>
      <td width="70" align="right">货品类别：</td>
      <td width="150"><label>
      <input type="text" name="nodename" id="nodename" size="16" value=<%=Request.Form("nodename")%>><a href="#" onClick="show_hide_div()"><img border="0" src="../img/choose.gif" width="21" height="17"></a></label><input type="hidden" name="typecode" value="<%=request.Form("typecode")%>" /></td>
      <td rowspan="3"><input type="submit" name="Submit" value=" 查 询 " onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  class="button" /></td>
    </tr>
    <tr>
      <td width="70" align="right">货品名称：</td>
      <td width="150">
	    <input type="hidden" name="goodscode" id="goodscode" value="<%=request.Form("goodscode")%>"/>
        <input type="text" name="goodsname" id="goodsname" size="16" value="<%=request.Form("goodsname")%>"/><a href="#" onClick="window.open ('selectgoods.asp', 'newwindow', 'height=600, width=800,top=100,left=150, toolbar=no,menubar=no, scrollbars=yes, resizable=no,location=no, status=no')"><img border="0" src="../img/choose.gif" width="21" height="17"></a></td>
      <td width="70" align="right">截止日期：</td>
      <td width="150"><input type="text" name="date" size="16" readonly value=<%
If Request("date") = "" Then
    Response.Write formatdate(date)
Else
    Response.Write Request("date")
End If
%>><%showdate("date")%></td>
    </tr>
    <tr>
      <td width="70" align="right">货品规格：</td>
      <td width="150">
        <input type="text" name="goodsunit" id="goodsunit" size="16" value="<%=request.Form("goodsunit")%>"/></td>
    
    </tr>
	<tr>
	   <td width="70" align="right">货品编码：</td>
	   <td width="150"><input type="text" name="code" value="<%=Request.Form("code")%>" ></td>
	</tr>
  </table>
<hr>
</form>
<%
sql = "delete from t_viewdepot"
conn.Execute(sql)
If Request("date") = "" Then
    s_date = formatdate(date)
Else
    s_date = Request("date")
End If

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

if Request("depot") = "" then
  s_depot = " and depotname in ("&s_depotpower&"'a')"
else
  s_depot = " and depotname = '" & Request("depot") & "'"
end if

If Request.Form("goodsname") = "" Then
    s_goodsname = ""
Else
    s_goodsname = " and  goodsname like '%"&Request("goodsname")&"%'"
End If

If Request.Form("code") = "" Then
    s_goodscode = ""
Else
    s_goodscode = " and  goodscode like '%"&Request.Form("code")&"%'"
End If

if request("nodename") = "" then
  s_type = ""
else
  s_type = " and code like '" & request("typecode") & "%'"
end if

if Request("goodsunit") = "" then
	sGoodsunit = ""
else
	sGoodsunit = " and goodsunit like '%" & Request("goodsunit") & "%'"
end if

if Request.Form("showzero") <> "showzero" then
	sShowzero = " and t_num > 0"
else
    sShowzero = ""
end if
sql = "select CONCAT('<a href=# onClick=detail(''',goods.goodscode,''')>显示明细</a>') as action,goods.goodscode,goods.goodsname,goods.goodsunit,goods.units,avgprice,t_num,avgprice*t_num as t_mon,goodstype from (select * from t_goods where 1=1"&s_goodsname&s_type&s_goodscode&sGoodsunit&") as goods left join (select cast((case when sum(qty) is null or sum(qty) = 0 then 0 else sum(qty*price)/sum(qty) end) as float) as avgprice,cast(sum(qty) as float) as t_num,goodscode from t_stock where 1=1"&s_depot&" group by goodscode) as s on goods.goodscode=s.goodscode where 1=1"


%>

<% call showpage(sql,"ViewDepot",1)
endconnection
%>
</td></tr></table>
</div>
</body>
</html>
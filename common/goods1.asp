<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<% call CheckAuthority("strBrowseGoods1") %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<title>货品资料</title>
<script>
$(document).ready(function() {
	$("#checkall").click(function() {
		if ($("#checkall").attr("checked") == true) {
				$(".c").attr("checked",true);
				var aa = document.getElementsByName("t");
				for (var j=0; j<aa.length; j++)
				$("#tt").val($("#tt").val() + aa[j].value + ",");
			}else{
				$(".c").attr("checked",false);
				$("#tt").val('');
			}
	});
	
	$(".c").click(function() {
		if ($(this).attr("checked") == true){
		$("#tt").val($("#tt").val() + $(this).val() + ",");
		}else{
		$("#tt").val($("#tt").val().replace($(this).val() + ",",""));
		}
	});
});
function changegoodstype(){
  if (document.getElementById('nodename').value == "") {
	 document.getElementById('typecode').value ="";
	  }	
}
function CreateReturnValue(){
　　 var nodecode=document.getElementById("goodstype");
     nodecode.value=document.getElementById('nodename').value; 
}
function del(){
if (document.getElementById('temp').value == "" ){
alert("请选贼要删除的货品资料！");
return false;
}
else{
if(!confirm('确定要删除该货品的资料吗?')){
return false
}else{
window.location.href='../action/deletegoods.asp?goodsid='+document.getElementById('temp').value
}
}
}
function edit(){
if(document.getElementById('temp').value == "" ){
alert("请选择要编辑的货品资料！");
return false;
}
}
</script>
</head>
<body style="background:#FFFFFF">
<div style="background:url(../images/main-02.gif); width:100%; height:20; padding-left:0" ><img src="../images/main-01.gif"></div>
<div style="padding-left:10px">
<table align="center"><tr><td>
<form name="form1" method="post" action="goods1.asp">
<table border="0">
<tr>
	<td align="right">货品编码：</td>
    <td><input type="text" id="goodscode" name="goodscode" size="16" value="<%=request.Form("goodscode")%>"></td>
    <td align="right">货品规格：</td>
    <td><input type="text" name="goodsunit" size="16" value="<%=request.Form("goodsunit")%>"></td>
    <td rowspan="2"><input type="submit" class="button"  onMouseOver="this.className='button_over';" onMouseOut="this.className='button';" value=" 查 询 "></td>
</tr>
<tr>
	<td align="right">货品名称：</td>
    <td><input type="text" name="goodsname" size="16" value="<%=request.Form("goodsname")%>"></td>
	<td align="right"></td>
    <td></td>
</tr>
</table>
</form>
请选择要修改的货品：
<tr valign="top">
<td>
<%
if Request.Form("goodscode") = "" then
	s_goodscode = ""
else
	s_goodscode = " and goodscode like '%"& Request.Form("goodscode") &"%'"
end if


if Request.Form("goodsname") = "" then
	s_goodsname = ""
else
	s_goodsname = " and goodsname like '%"& Request.Form("goodsname") &"%'"
end if

if Request.Form("goodsunit") = "" then
	s_goodsunit = ""
else
	s_goodsunit = " and goodsunit like '%"& Request.Form("goodsunit") &"%'"
end if

sql = "select CONCAT('<input type=checkbox name=t class=c value=''',goodscode,'''>') as action,t_goods.* from t_goods where 1=1 " & s_goodscode & s_goodsname & s_goodsunit
'response.write(sql)
call showpage(sql,"goods1",1)

endconnection
 %>
请输入要修改的类别：
<form name="changetype" action="../action/changetype.asp" method="post">
<input type="hidden" size="60" id="tt" name="tt">
货品类别：<input type="text" id="nodename" name="nodename" size="16" value="" onpropertychange="changegoodstype();"><input type="hidden" name="typecode" value=""><a href="#" onClick="window.open ('tree.asp?type=goods', 'newwindow', 'height=400, width=200,toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=no')"><img border="0" src="../img/choose.gif" width="21" height="17"></a>
<input type="submit" class="button"  onMouseOver="this.className='button_over';" onMouseOut="this.className='button';" value="批量修改货品类别">
</form>
</td></tr></table>
</div>
</body>
</html>

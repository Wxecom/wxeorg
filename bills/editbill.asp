<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<%
if instr(Request.QueryString("billcode"),"CG") <> 0 then
	Power = Authority("strEditCG")
end if
if instr(Request.QueryString("billcode"),"XS") <> 0 then
	Power = Authority("strEditXS")
end if
if instr(Request.QueryString("billcode"),"CT") <> 0 then
	Power = Authority("strEditCT")
end if
if instr(Request.QueryString("billcode"),"XT") <> 0 then
	Power = Authority("strEditXT")
end if
if instr(Request.QueryString("billcode"),"CD") <> 0 then
	Power = Authority("strEditCD")
end if
if instr(Request.QueryString("billcode"),"XD") <> 0 then
	Power = Authority("strEditXD")
end if
if instr(Request.QueryString("billcode"),"RK") <> 0 then
	Power = Authority("strEditIn")
end if
if instr(Request.QueryString("billcode"),"CK") <> 0 then
	Power = Authority("strEditOut")
end if
if instr(Request.QueryString("billcode"),"DB") <> 0 then
	Power = Authority("strEditDB")
end if
if instr(Request.QueryString("billcode"),"PD") <> 0 then
	Power = Authority("strEditPD")
end if
if instr(Request.QueryString("billcode"),"LL") <> 0 then
	Power = Authority("strEditLL")
end if
if instr(Request.QueryString("billcode"),"TL") <> 0 then
	Power = Authority("strEditTL")
end if
s_type = left(request.QueryString("billcode"),2)
sql = "select * from t_bill where billcode='"&Request.QueryString("billcode")&"'"
Set rsBill = server.CreateObject("adodb.recordset")
rsBill.Open sql, conn, 1, 1
set rsBillType = Server.CreateObject("adodb.recordset")
sql = "select * from dict_bill where name = '"& rsBill("billtype") &"'"
rsBillType.open sql, conn, 1, 1
%>
<HTML><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<link rel="stylesheet" href="../style.css" type="text/css">
<META http-equiv=Pragma content=no-cache>
<script language="JavaScript" src="../js/jquery.js"></script>
<script language="JavaScript" src="../js/addrow.js"></script>
<script>
function bar(){
  if ((event.keyCode == 13) && ($("#barcode").val()!="")){
	
    $.post("barcode.asp",{date:escape($("#date").val()),barcode:escape($("#barcode").val()),depot:escape($("#depot").val())},
	
	function(data)
	{ 
	    if(data!="")
		{
	   addrow_bar();
	   i=$("#rowcount").val();
	   str=data.split("|");
	   $("#goodscode"+i).val(str[0]);
	   $("#goodsname"+i).val(str[1]);
	   $("#goodsunit"+i).val(str[2]);
	   $("#units"+i).val(str[3]);
	   $("#number"+i).val("1");
	   $("#aveprice"+i).val(str[8]);
	   $("#fact_num"+i).val(str[7]);
	   $("#remark"+i).val(str[6]);
	   $("#number"+i).addClass("number");
	    $("#money"+i).addClass("money");
      

	   <%if a("billtype") = "采购订货" then%>
	     $("#price"+i).val(str[4]);
	   <%elseif a("billtype") = "销售订货" then%>
	     $("#price"+i).val(str[5]);
	   <%elseif a("billtype") = "采购入库" then%>
	     $("#price"+i).val(str[4]);
	   <%elseif a("billtype") = "销售出库" then%>
	     $("#price"+i).val(str[5]);
	   <%elseif a("billtype") = "其他入库" then%>
	     $("#price"+i).val(str[4]);
	   <%elseif a("billtype") = "其他出库" then%>
	     $("#price"+i).val(str[5]);
	   <%elseif a("billtype") = "仓库盘点" then%>
	     $("#price"+i).val(str[8]);
	   <%elseif a("billtype") = "仓库调拨" then%>
	     $("#price"+i).val(str[8]);
	   <%end if%>
	   }
	   else
	   {
	    alert('没有相关数据');
	   }
	}
	);
	$("#barcode").val("");
  }
}

function check(){
	
if(document.sample.cust.value=="")
{
alert("请选择往来单位！");
document.sample.cust.focus();
return false;
}
if (document.getElementById("sqd").rows.length == 1)
{
alert("请添加货品！");
return false;
}
}
function edit_row(i)
{
	openwin('../common/selectgoods.asp?depot='+document.getElementById("depot").value+'&type=<%=s_type%>&adddate='+document.all.date.value+'&id='+i, 800,600)
}

function add_row(){
	if ($("#depot").val() == "")
	{
	alert("请选择仓库！");
	document.sample.depot.focus();
	return false;
	}
	openwin('../common/selectgoods.asp?depot='+document.getElementById("depot").value+'&type=<%=s_type%>&adddate='+document.all.date.value,800,600)
}
function counttotal(){

var vTotalNumber=0;
var vTotalMoney=0;//总金额的初始值为0;
$(".money").each( 
	function(i)
	{
	var vTempValue=$(this).val();
		if(vTempValue=="")
		{
			vTempValue=0;
		}
	vTotalMoney=vTotalMoney+parseFloat(vTempValue);//计算总费用
	document.all.zdprice.value=vTotalMoney;
	document.all.yfprice.value=parseFloat((document.all('zdprice').value))-parseFloat((document.all('zkprice').value));
	}
)//遍历结束
$("#total_money").val(vTotalMoney); //将总费用显示到对应文本框对象中
$(".number").each( 
	function(i)
	{
	var vTempValue=$(this).val();
		if(vTempValue=="")
		{
			vTempValue=0;
		}
	vTotalNumber=vTotalNumber+parseFloat(vTempValue);//计算总费用
	}
)//遍历结束
$("#total_number").val(vTotalNumber); //将总费用显示到对应文本框对象中
}
function count(i){
r_price=Number(document.getElementById("price"+i).value);

r_number=Number(document.getElementById("number"+i).value);
document.getElementById("money"+i).value=(r_price*r_number).toFixed(2);
document.getElementById('total_money').value=0;
document.getElementById('total_number').value=0;
counttotal();
}
function del_row(i) {
	document.getElementById("sqd").deleteRow(i);
	$("#rowcount").val($("#rowcount").val() - 1);
	counttotal();	
}
</script>
<script language="javascript">
function changesk(){
	document.all.yfprice.value=parseFloat((document.all('zdprice').value))-parseFloat((document.all('zkprice').value))
}
function doPrint(billcode) { 
openwin('../action/print.asp?billcode='+billcode,800,600)
}
</script>
</HEAD>
<BODY>
<div align="left" style="margin-left:10px">
  <input name="button" type="button" class="button but_mar" onClick="doPrint('<%=rsbill("billcode")%>')" value=" 打 印 ">
</div>
<%if (rsBill("billtype")="采购退货" or rsBill("billtype")="销售退货") then%>
<form method="POST" onKeyPress="if(event.keyCode==13){return false;}" action="../action/savebackbill.asp?type=<%=a("billtype")%>&add=false" name="sample">
<%elseif rsBill("billtype")="库存盘点" then%>
<form method="POST" onKeyPress="if(event.keyCode==13){return false;}" action="../action/savepandian.asp?type=<%=a("billtype")%>&add=false" name="sample">
<%elseif rsBill("billtype")="仓库调拨" then%>
<form method="POST" onKeyPress="if(event.keyCode==13){return false;}" action="../action/savediaobo.asp?type=<%=a("billtype")%>&add=false" name="sample">
<%else%>
<form method="POST" onKeyPress="if(event.keyCode==13){return false;}" action="../action/savebill.asp?type=<%=a("billtype")%>&add=false" name="sample">
<%end if%>
<table align="center" border="0">
<tr>
<td>
<br>
<div align="center"><span class="style1"><%
Response.Write rsBillType("caption")
%></span>
<%
if rsBill("check") = True then
	Response.Write "<span class='checkstyle'>[已审核]</span>"
else
	Response.Write "<span class='Uncheckstyle'>[未审核]</span>"
end if
%>
</div>
<div align="center"><span class="style1"><%=Request("billcode")%></span></div><br>
<INPUT type=hidden name="billcode" value=<%=Request("billcode")%>>
<input type="hidden" name="_check" value="<%=rsBill("check")%>">
<table id="tbl" style="font-size:12px" width="860" border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFEE">
<tbody>
<tr>
  <td align="right" width="70">制单日期：</td>
  <td width="320" colspan="3"><input type="text" name="date" id="date" size="16" value=<%=rsBill("adddate")%> readonly><%showdate("date")%><font color="red">*</font></td>
  <td align="right" width="70">
<%
Response.Write rsBillType("f_dw")
%></td>
  <td width="320" colspan="3">
<%if rsBill("billtype") = "库存盘点" then%>
<%elseif rsBill("billtype") = "仓库调拨" then%>
<%call showdepot("cust",rsBill("custname"))%><font color="red">*</font>
<%elseif rsBill("billtype") = "领料出库" then
call ShowCombo("t_department","name","cust",rsBill("custname"))
elseif rsBill("billtype") = "退料入库" then
call ShowCombo("t_department","name","cust",rsBill("custname"))
else%>
<input type="text" name="Cust" size="35" id="cust"  value=<%=rsBill("custname")%>><a href="#" onClick="JavaScript:window.open ('../common/selectcust.asp', 'newwindow', 'height=600, width=800,top=100,left=150, toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=no')"><img border="0" src="../img/choose.gif" width="21" height="17"></a><font color="red"></font>
<%end if%>
</td></tr>
<tr>
  <td  align="right">
  <%if rsBill("billtype")="仓库调拨" then%>
  调&nbsp;出&nbsp;仓：
  <%else%>
  仓&nbsp;&nbsp;&nbsp;&nbsp;库：
  <%end if%>
  </td>
  <td colspan="3"><%call showdepot("depot",rsBill("depotname"))%><font color="red">*</font></td>
		
 
  	<td align="right">经&nbsp;办&nbsp;人：</td><td><select  style="width:118" name="user">
    <%
   sql = "select * from t_Employee"
   Set rs = conn.Execute(sql)
   Do While rs.EOF = False
    s_name = rs("name")
    If s_name = rsBill("username") Then
        Response.Write "<option value="&s_name&" selected>"&s_name&"</option>"
    Else
        Response.Write "<option value="&s_name&">"&s_name&"</option>"
    End If
    rs.movenext
   Loop
   close_rs(rs)%></select><font color="red">*</font></td>
   <%if rsBill("billtype") = "采购入库" then%>
   <td align="right">帐&nbsp;&nbsp;&nbsp;&nbsp;户：</td>
  <td><select name="account" style="width:118"><option value=""></option>
               <%
	           sql_account = "select * from t_account"
	    	       Set rs_account = server.CreateObject("adodb.recordset")
				  rs_account.open sql_account,conn,1,1
				  Do While rs_account.EOF = False
				   if trim(rsBill("account"))=trim(rs_account("account")) then
	 			   Response.Write "<option value="&rs_account("account")&" selected>"&rs_account("account")&"</option>"
				   else
				   Response.Write "<option value="&rs_account("account")&">"&rs_account("account")&"</option>"
				   end if
				  rs_account.movenext
				  loop
				  close_rs(rs_account)
				  %></select></td>
  <%end if%>
  <%if rsBill("billtype") = "销售出库" then%>
    <td align="right">帐&nbsp;&nbsp;&nbsp;&nbsp;户：</td>
  <td><select name="account"><option value=""></option>
               <%
	           sql_account = "select * from t_account"
	    	       Set rs_account = server.CreateObject("adodb.recordset")
				  rs_account.open sql_account,conn,1,1
				  Do While rs_account.EOF = False
				   if trim(rsBill("account"))=trim(rs_account("account")) then
	 			   Response.Write "<option value="&rs_account("account")&" selected>"&rs_account("account")&"</option>"
				   else
				   Response.Write "<option value="&rs_account("account")&">"&rs_account("account")&"</option>"
				   end if
				  rs_account.movenext
				  loop
				  close_rs(rs_account)
				  %></select></td>
  <%end if%>
<tr>
<td align="right">备&nbsp;&nbsp;&nbsp;&nbsp;注：</td><td colspan="5" ><input type="text" name="memo" size="78" value="<%=rsBill("memo")%>"></td>
<td align="right">制&nbsp;单&nbsp;人：</td><td><input type="maker" size="16" readonly value="<%=rsBill("maker")%>"><font color="red">*</font></td>
</tr>
<% if  (rsBill("billtype") = "采购入库") then%>
<tr>
  <td align="right">整单金额：</td>
  <td><input type="text" name="zdprice" size="16" onKeyUp=" changesk()" value="<%=rsBill("zdprice")%>"></td>
  <td align="right" width="79" >折扣金额：</td>
  <td><input type="text" name="zkprice" size="15" onKeyUp=" changesk()" value="<%=rsBill("zkprice")%>"></td>
 
  <td align="right">应付金额：</td>
  <td><input type="text" name="yfprice" size="15" value="<%=rsBill("yfprice")%>"></td>
  <td align="right">实付金额：</td>
  <td><input type="text" name="pay" size="15" value="<%=rsBill("pay")%>"></td>
  
</tr>
<%end if%>
<% if  (rsBill("billtype") = "销售出库") then%>
<tr>
  <td align="right">整单金额：</td>
  <td ><input type="text" name="zdprice" size="16"  onKeyUp=" changesk()" value="<%=rsBill("zdprice")%>"></td>
  <td  width="79" align="right">折扣金额：</td>
  <td><input type="text" name="zkprice" size="16" onKeyUp=" changesk()"  value="<%=rsBill("zkprice")%>"></td>
  

  <td align="right">应收金额：</td>
  <td><input type="text" name="yfprice" size="15" value="<%=rsBill("yfprice")%>"></td>
  <td align="right">实收金额：</td>
  <td><input type="text" name="pay" size="15" value="<%=rsBill("pay")%>"></td>
  
</tr>
<%end if%>

</tbody>

<table>
 <tr>
   
	  <td valign="top" width="65">
	<input name="add" type="button" onClick=add_row() class="button" value="添加商品 ">
	</td>
	  <td valign="top" width="280">
	&nbsp;&nbsp;条码录入框：<input name="barcode" id="barcode" size="16" onKeyUp="bar();">
	 <td>
	</tr>
	</table>
</table>
 <td valign="top">
 

<table style="font-size:12px" width="860" id="sqd"  align="center" border="1" cellpadding="2" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFEE">
<tr align="center">
  <th width=3%>&nbsp;</th>
  <th width=5%>删除</th>
  <th width=10%>商品编码</th>
  <th width=15%>商品名称</th>
  <th width=10%>商品规格</th>
  <th width=7%>单位</th>
  <th width=10%>单价</th>
  <th width=10%>数量</th>
  <th width=10%>金额</th>
  <th width=20%>备注</th>

</tr>
<%
sql = "select * from t_billdetail where billcode='"&rsBill("billcode")&"'"
Set rs_detail = server.CreateObject("adodb.recordset")
rs_detail.Open sql, conn, 1, 1
%>
<input type="hidden" id="rowcount" name="rowcount" value=<%=rs_detail.recordcount%>>
<%
For i = 1 To rs_detail.recordcount 
%>
<tr>
<td><%=i%></td>
<td><a href='#' onClick='del_row(<%=i%>)'>删除</a></td>
<td><input name=goodscode type=text id=goodscode<%=i%> size=8 value="<%=rs_detail("goodscode")%>" onClick="edit_row(<%=i%>)"></td>
<td><input name=goodsname type=text id=goodsname<%=i%>  size=16 value="<%=rs_detail("goodsname")%>" onClick="edit_row(<%=i%>)"></td>
<td><input name=goodsunit type=text id=goodsunit<%=i%>  size=8 value="<%=rs_detail("goodsunit")%>"></td>
<td><input name=units type=text id=units<%=i%>  size=3 value="<%=rs_detail("units")%>"></td>
<td><input onkeyup=count(<%=i%>) name=price type=text id=price<%=i%>  size=8 value="<%=showprice(rs_detail("price"))%>"></td>
<td><input onkeyup=count(<%=i%>) name=number type=text class='number' id=number<%=i%>  size=8 value="<%=shownumber(rs_detail("number"))%>"></td>
<td><input name=money type=text id=money<%=i%> size=8 class='money' readonly value="<%=showprice(rs_detail("money"))%>"></td>
<td><input name=remark type=text id=remark<%=i%> size=20 value="<%=rs_detail("DetailNote")%>">
<input name=aveprice type=hidden id=aveprice<%=i%> value="<%=rs_detail("inprice")%>"></td>
</tr>
<%
total_number = total_number + cdbl(rs_detail("number"))
total_money  = total_money  + cdbl(rs_detail("money"))
total_money  = (formatnumber(total_money,2,-1,false,false))
rs_detail.movenext
Next
%>
</table>
<table style="font-size:12px" width="860" id="sqd1" align="center" border="1" cellpadding="2" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFEE">
<tr align="center">
  <th width=60%>合计</th>
  <th width=10%><input id="total_number" type="text" size="8" value="<%=total_number%>"></th>
  <th width=10%><input id="total_money" type="text" size="8" value="<%=total_money%>"></th>
  <th width=20%></th>
</tr>
</table>
<%
if power = "True" then
if rsBill("check") <> True then%>
<div align="center" class="but_mar"><input type="submit" onClick="return check();" class="button" value=" 保 存 " name="B1"></div>
<%
end if
end if
%>

</td></tr></table>
<%
close_rs(a)
close_rs(rs_detail)
endconnection
%>
</form>
</BODY>
</HTML>

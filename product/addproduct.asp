<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<% call CheckAuthority("strAddZZ") %>
<HTML><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<link rel="stylesheet" href="../style.css" type="text/css">
<META http-equiv=Pragma content=no-cache>
<script language="JavaScript" src="../js/addrow.js"></script>
<script language="JavaScript" src="../js/jquery.js"></script>
<script>
function setnum(){
	for (k=1;k<document.all.sqd.rows.length;k++){
		$("#number"+k).attr("value",Number($("#modelnumber"+k).val())*Number($("#totalnumber").val())/Number($("#totalnumber1").val()));
	}
}



function checkbill(){
var temp = "";
var temp1 = "";
$.post("checkbillcode_ajax.asp",{billcode:escape($("#billcode").val())},
function(data)
{ 
   if (data == "False"){
	   alert("有重复单号，请再次点击保存");
		temp = $("#tittle_lable").html();
		temp1 = "000"+String(Number(temp.substring(14))+1);
	   $('#tittle_lable').html(temp.substring(0,14)+temp1.substring(temp1.length-4)); 
	   $('#billcode_input').attr("value",$('#tittle_lable').html());
	   return false;
   }
}
);		
	}
	
function bar(){
  if ((event.keyCode == 13) && ($("#barcode").val()!="")){
	addrow_bar();
    $.post("barcode.asp",{date:escape($("#date").val()),barcode:escape($("#barcode").val()),depot:escape($("#depot").val())},
	function(data)
	{ 
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
	   $("#price"+i).val(str[4]);
	}
	);
	$("#barcode").val("");
  }
}
function check(){

if (document.sample.depot1.value=="")
{
alert("请选择出库仓！");
document.sample.depot1.focus();
return false;
}

if (document.sample.depot2.value=="")
{
alert("请选择入仓库！");
document.sample.depot2.focus();
return false;
}

if (document.getElementById("sqd").rows.length == 1)
{
alert("请添加货品！");
return false;
}
for (i=1;i<=Number(document.getElementById("rowcount").value);i++){
  need_number = Number(document.getElementById("number"+i).value);
  fact_number = Number(document.getElementById("fact_num"+i).value);
  <%
  sqlFCK = "select fuchuku from t_softinfo"
  set rsFCK = Server.CreateObject("adodb.recordset")
  rsFCK.open sqlFCK,conn,1,1
  if rsFCK("fuchuku") = false then
  %>
  if (need_number>fact_number){
  alert(document.getElementById("goodsname"+i).value+"库存数量不足！");
  return false;
  }
  <%end if%>
}
}
function edit_row(i)
{
	window.open ('../common/selectgoods.asp?depot='+document.getElementById("depot").value+'&type=<%=request.QueryString("type")%>&adddate='+document.all.date.value+'&id='+i, 'newwindow', 'top=100,left=150,height=600, width=800,toolbar=no,menubar=no, scrollbars=yes, resizable=no,location=no, status=no')
}

function totalcount(){

}

function count(i){
if (i!= 0){
r_price=Number(document.getElementById("price"+i).value);
r_number=Number(document.getElementById("number"+i).value);
document.getElementById("money"+i).value=(r_price*r_number).toFixed(2);
}document.getElementById('total_money').value=0;
document.getElementById('total_number').value=0;
    for (j=1;j<document.all.sqd.rows.length;j++)
	{	
	document.getElementById('total_money').value=(Number(document.getElementById('total_money').value)+Number(document.getElementById("money"+j).value)).toFixed(2);
	document.getElementById('total_number').value=Number(document.getElementById('total_number').value)+Number(document.getElementById("number"+j).value);
	}
	document.getElementById('totalmoney').value=document.getElementById('total_money').value;
	document.getElementById('totalprice').value=(Number(document.getElementById('totalmoney').value)/Number(document.getElementById("totalnumber").value)).toFixed(2);
}

var j=0
function add_row(){
	if ($("#depot").val() == "")
	{
	alert("请选择仓库！");
	document.sample.depot.focus();
	return false;
	}
    //addrow();
	window.open ('../common/selectgoods.asp?depot='+document.getElementById("depot1").value+'&type=<%=request.QueryString("type")%>&adddate='+document.all.date.value+'&id='+j, 'newwindow', 'top=100,left=150,height=600, width=800,toolbar=no,menubar=no, scrollbars=yes, resizable=no,location=no, status=no')
}
function del_row() {
    if(sqd.rows.length==1) return;
    var checkit = false
    for (var i=0;i<document.all.delNumber.length;i++) {
            if (document.all.delNumber[i].checked) {
            checkit=true;
            sqd.deleteRow(i+1)
			count(i)
            break;
            }
    }
    if (checkit) {
        for(i=1;i<sqd.rows.length;i++){
        sqd.rows[i].cells[0].innerText=i
		document.getElementById("rowcount").value=sqd.rows.length-1
        }

    } else
    {
    alert("请选择一个要删除的对象");
    return false}
}

function selectmodel(){
	window.open ('selectmodel.asp?depot1='+document.getElementById("depot1").value+'&depot2='+document.getElementById("depot2").value+'&username='+document.getElementById("user").value+'&adddate='+document.all.date.value, 'selectmodel', 'top=100,left=150,height=600, width=800,toolbar=no,menubar=no, scrollbars=yes, resizable=no,location=no, status=no')	
}
function addmodel(){
	window.open ('../model/addmodel.asp?add=true', 'main')	
}
</script>
</HEAD>
<BODY>
<%
sql_bill = "select * from s_model where billcode = '" & Request.QueryString("billcode") & "'"
set rs_bill = server.CreateObject("adodb.recordset")
rs_bill.open sql_bill,conn,1,1
if rs_bill.recordcount > 0 then
  sGoodsCode = rs_bill("totalgoodscode")
  sGoodsName = rs_bill("totalgoodsname")
  sGoodsUnit = rs_bill("totalgoodsunit")
  sUnits     = rs_bill("totalunits")
  sNumber    = rs_bill("totalnumber")
  sPrice     = "0"
  sMoney     = "0.00"
  sRemark    = ""
else
  sGoodsCode = ""
  sGoodsName = ""
  sGoodsUnit = ""
  sUnits     = ""
  sNumber    = ""
  sPrice     = "0"
  sMoney     = "0.00"
  sRemark    = ""
end if
%>
<table align="center" border="0"><tr><td>
<form method="POST" onKeyPress="if(event.keyCode==13){return false;}" action="saveproduct.asp" name="sample">

<br><div align="center"><span class="style1">产品入库单</span></div>
<div align="center"><span class="style1"><%
today = Now()
tdate = Year(today) & "-" & Right("0" & Month(today), 2) & "-" & Right("0" & Day(today), 2)
sql = "select * from t_bill where diaobo=0 and billcode like 'PR-"&tdate&"%' order by billcode desc"
Set rs = server.CreateObject("adodb.recordset")
rs.Open sql, conn, 1, 1
If rs.recordcount = 0 Then
    s_billcode = "PR-"&tdate&"-0001"
Else
    s_temp = rs("billcode")
    s_temp = Right(s_temp, 4) + 1
    s_billcode = "PR-"&tdate&"-"&Right("000"&s_temp, 4)
End If
close_rs(rs)
Response.Write "<label id=tittle_lable>"&s_billcode&"</label>"
Response.Write "<INPUT type=hidden id=""billcode_input"" name=""billcode"" value="&s_billcode&">"
%></span></div><br>
<table id="tbl" style="font-size:12px" width="780" border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFEE">
<tbody>
<tr>
<td align="right" width="70">制单日期：</td>
<td width="320"><input type="text" name="date" id="date" size="16" readonly value=<%=formatdate(date)%>><%showdate("date")%><font color="red">*</font></td>
<td align="right" width="70">出&nbsp;库&nbsp;仓：</td>
<td width="320"><%call showdepot("depot1",request.QueryString("depot1"))%></td>
</tr>
<tr><td align="right">入&nbsp;库&nbsp;仓：</td>
    <td><%call showdepot("depot2",request.QueryString("depot2"))%><font color="red">*</font></td>
	<td align="right">制&nbsp;单&nbsp;人：</td>
	<td><input type="text" size="16" name="maker" readonly value="<%=request.cookies("username")%>"><font color="red">*</font></td></tr>
  <tr><td align="right">备&nbsp;&nbsp;&nbsp;&nbsp;注：</td><td><input onKeyDown="" type="text" name="memo" size="30"></td>
  <td align="right">经&nbsp;办&nbsp;人：</td>
  <td><%call ShowCombo("t_employee","name","user",request.QueryString("username"))%><font color="red">*</font></td>
  </tr></tbody></table>
<table border="0" align="center">
<tr>
	<th>产品编码</th>
    <th>产品名称</th>
    <th>产品规格</th>
    <th>单位</th>
    <th>产品数量</th>
    <th>产品价格</th>
    <th>产品金额</th>
    <th>备注</th>
</tr>
<tr>
	<td><input type="text" id="totalgoodscode" name="totalgoodscode" size="10" value="<%=sGoodsCode%>"><a href="#" onClick="window.open ('selectgoods.asp', 'selectgoods', 'height=600, width=800,toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=no')"><img border="0" src="../img/choose.gif" width="21" height="17"></td>
    <td><input type="text" id="totalgoodsname" name="totalgoodsname" size="20" value="<%=sGoodsName%>"></td>
    <td><input type="text" id="totalgoodsunit" name="totalgoodsunit" size="18" value="<%=sGoodsUnit%>"></td>
    <td><input type="text" id="totalunits" name="totalunits" size="5" value="<%=sUnits%>"></td>
    <td><input type="text" id="totalnumber" name="totalnumber" onpropertychange="totalcount();" size="10" value="<%=sNumber%>">
    	<input type="hidden" id="totalnumber1" name="totalnumber1" value="<%=sNumber%>"></td>
    <td><input type="text" id="totalprice" name="totalprice" size="10" value="0"></td>
    <td><input type="text" id="totalmoney" name="totalmoney" size="10" value="0.00"></td>
    <td><input type="text" id="totalremark" name="totalremark" size="10" value="">
    	<input type="hidden" id="totalaveprice" name="totalaveprice" size="10" value="<%=sNumber%>"></td>
</tr>
</table>
<div align="left" class="but_mar">
    <input name="add" type="button" onClick=add_row() class="button" value=" 添 加 ">
	<input name="del" onclick=del_row() type="button" class="button" value=" 删 除 ">
    <input name="model" onClick="selectmodel()" type="button" class="button" value="选择模板">
    <input name="model" onClick="addmodel()" type="button" class="button" value="新增模板">
    <input name="countnumberdetail" onClick="setnum();" type="button" class="button" value="&nbsp;计&nbsp;&nbsp;算&nbsp;">
&nbsp;&nbsp;条码录入框：<input name="barcode" id="barcode" size="16" onKeyUp="bar();"></div>
<table style="font-size:12px" width="780" id="sqd" align="center" border="1" cellpadding="2" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFEE">
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
Response.Write "<input type='hidden' id='rowcount' name='rowcount' value="&rs_bill.recordcount&">"
if rs_bill.recordcount > 0 then
	for i = 1 to rs_bill.recordcount
		if Request.QueryString("adddata") = "" then
			sDate = ""
		else
			sDate = " and adddate<='"&Request.QueryString("adddate")&"'"
		end if
		if Request.QueryString("depot1") = "" then
		  sDepot = ""
		else
		  sDepot = " and depotname = '"&request.QueryString("depot1")&"'"
		end if
		sql_detail = "select s1.goodscode,goodsname,goodsunit,units,inprice,outprice,avgprice,detailnumber,detailnumber*avgprice as [money] from (select goodscode,goodsname,goodsunit,units,inprice,outprice,remark from t_goods where goodscode = '" & rs_bill("detailgoodscode") & "') as s1 left join (select goodscode,ifnull(case when round(sum(number*flag),9)=0 then 0 else round(round(sum(inprice*number*flag),9)/round(sum(number*flag),9),9) end,0) as avgprice from s_count where goodscode = '" & rs_bill("detailgoodscode") & "'" & sDate & sDepot & " group by goodscode) as s2 on s2.goodscode = s1.goodscode left join (select detailgoodscode,detailnumber from s_model where billcode = '" & Request.QueryString("billcode") & "' and detailgoodscode = '" & rs_bill("detailgoodscode") & "') as s3 on s3.detailgoodscode=s1.goodscode"
		set rs_detail = server.CreateObject("adodb.recordset")
		rs_detail.Open sql_detail, conn, 1, 1
		Response.Write "<tr>"
		Response.Write "<td>"&i&"</td>"
		Response.Write "<td><input name=delNumber type=radio></td>"
		Response.Write "<td><input name='goodscode"&i&"' type=text class=INPUT1 id='goodscode"&i&"' size=8 onClick='edit_row("&i&")' value="&rs_detail("goodscode")&"></td>"
		Response.Write "<td><input name=goodsname"&i&" type=text class=INPUT1 id=goodsname"&i&"  size=16  onClick=edit_row("&i&") value='" & rs_detail("goodsname") & "'></td>"
		Response.Write "<td><input name=goodsunit"&i&" type=text class=INPUT1 id=goodsunit"&i&"  size=8 value='" & rs_detail("goodsunit") & "'></td>"
		Response.Write "<td><input name=units"&i&" type=text class=INPUT1 id=units"&i&"  size=3 value='" & rs_detail("units") & "'></td>"
		Response.Write "<td><input onpropertychange='count("&i&")' name=price"&i&" type=text class=INPUT1 id=price"&i&"  size=8 value='" & rs_detail("avgprice") & "'></td>"
		Response.Write "<td><input onpropertychange='count("&i&")' name=number"&i&" type=text class=INPUT1 id=number"&i&"  size=8 value='" & rs_detail("detailnumber") & "'></td>"
		Response.Write "<td><input name=money"&i&" type=text class=INPUT1 id=money"&i&"  size=8 value='" & rs_detail("money") & "'></td>"
		Response.Write "<td><input name=remark"&i&" type=text class=INPUT1 id=remark"&i&"  size=20 value=''><input name=aveprice"&i&" type=hidden class=INPUT1 id=aveprice"&i&"  size=20 value='" & rs_detail("avgprice") & "'><input name=modelnumber"&i&" type=hidden class=INPUT1 id=modelnumber"&i&"  size=20 value='" & rs_detail("detailnumber") & "'></td>"
		Response.Write "</tr>"
		rs_bill.movenext
	next
end if
%>
</table>
<% response.Write(sql_detail) %>
<table style="font-size:12px" width="780" id="sqd1" align="center" border="1" cellpadding="2" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFEE">
<tr align="center">
  <th width=60%>合计</th>
  <th width=10%><input id="total_number" type="text" size="8"></th>
  <th width=10%><input id="total_money" type="text" size="8"></th>
  <th width=20%></th>
</tr>
</table>
<div align="center" class="but_mar"><input type="submit" onMouseDown="checkbill();" onClick="return check();" class="button" value=" 保 存 " name="B1"></div>
</form>
</td></tr></table>
<%
if Request.QueryString("billcode") <> "" then
if rs_detail.recordcount > 0 then%>
<script>count(1)</script>
<%
end if
end if
%>
<%endconnection%>
</BODY>
</HTML>

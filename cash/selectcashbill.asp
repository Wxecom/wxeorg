<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<html>
<head>
<script>
function selectbill(closeWindow,i){
	var bill=window.opener.document.getElementById("bill");
	var cust=window.opener.document.getElementById("cust");
	var mon=window.opener.document.getElementById("mon");
	bill.value=document.getElementById('tbl').rows[i].cells[1].innerHTML;
	cust.value=document.getElementById('tbl').rows[i].cells[3].innerHTML;
	mon.value=document.getElementById('tbl').rows[i].cells[5].innerHTML;
　if (closeWindow){
　　window.opener=null;
　　window.close();
　}
}

function CheckValue()
{
	alert('请填写制单日期！');
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<title>单据选择</title>
</head>
<body>
<%
if request.Form("date1") = "" then
s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
else
s_date1 = request.Form("date1")
end if
if request.Form("date2") = "" then
s_date2 = formatdate(date)
else
s_date2 = request.Form("date2")
end if
%>
<table width="100%" border="0">
  <tr>
    <td><form name="form1" method="post" action="selectcashbill.asp?type=<%=Request("type")%>">
      日期从
      <label>
      <input type="text" name="date1" size="16" value=<%=s_date1%>><%showdate("date1")%>
      </label>
     到
     <label>
     <input type="text" name="date2" size="16" value=<%=s_date2%>><%showdate("date2")%>
     </label>    
     <label>
     <input type="submit" name="Submit" onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  class="button" value=" 查 找 ">
     </label>
    </form>    </td>
  </tr>
</table>
      <table id="tbl" border="1" cellpadding="3" cellspacing="0" bgcolor="#000000" align="center" empty-cells:show width="100%">
      <tr bgcolor="#FFFFFF"  style="background-color:#999999;   border-bottom:2px   solid   black;   height:25px"   align="center">
	  <th>行号</th><th>单号</th><th>日期</th><th>供应商/客户名称</th><th>应付/收金额</th><th>未付/收金额</th></tr>   
	  <%
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

sql_cust = "select custname from t_user where username='"&request.cookies("username")&"'"
Set rs_cust = conn.Execute(sql_cust)
arr = split(rs_cust("custname"),",")
temp_cust = ""
if ubound(arr) <> -1 then
  for i = lbound(arr) to ubound(arr)-1
    temp_cust = temp_cust & "'" & arr(i) & "',"
  next
end if
sCustname = ""

If Request("type") = "FK" Then
    s_type = "采购入库"
Else
    s_type = "销售出库"
End If
if (Request.Form("date1")="") or (Request.Form("date2")="") then
  s_date = ""
else
  s_date = " and adddate>='"&request.Form("date1")&"' and adddate<='"&request.Form("date2")&"'"
end if

sql = "select adddate, custname, billcode, "
sql = sql + "yfprice - ifnull((select sum(yfprice) from t_bill where t_bill.check = 1 and planbillcode = bill.billcode), 0) as mon, "
sql = sql + "yfprice - ifnull((select sum(yfprice) from t_bill where t_bill.check = 1 and planbillcode = bill.billcode), 0) - ifnull((select sum(money) from t_cash where billcode = bill.billcode),0) - pay as wmon "
sql = sql + "from t_bill as bill where bill.check = 1 and billtype = '" &s_type& "' and adddate between '"& s_date1 &"' and '"& s_date2 &"'"
sql = sql + " and yfprice - ifnull((select sum(yfprice) from t_bill where t_bill.check = 1 and planbillcode = bill.billcode), 0) - ifnull((select sum(money) from t_cash where billcode = bill.billcode),0) - pay > 0"

Set rs = server.CreateObject("adodb.recordset")
rs.Open sql, conn, 1, 1

	sql_sys = "select * from T_SoftInfo"
	Set rs_sys = conn.Execute(sql_sys) 
  rs.pagesize = rs_sys("pagerecord")
  If request("page")<>"" Then
    epage = CInt(request("page"))
    If epage<1 Then epage = 1
      If epage>rs.pagecount Then epage = rs.pagecount
    Else
      epage = 1
    End If
  If Not rs.EOF Then rs.absolutepage = epage end if
    	
For i = 0 To rs.pagesize -1
  If rs.bof Or rs.EOF Then Exit For
%>
<tr bgcolor="#FFFFFF" align="center"  onDblClick="selectbill(true,<%=i+1%>);"   onMouseOver="over()"   onClick="change()"   onMouseOut="out()">
  <td><%=i+1%></td>
  <td><%=rs("billcode")%></td>
  <td><%=rs("adddate")%></td>
  <td><%=rs("custname")%></td>
  <td><%=rs("mon")%></td>
  <td><%=rs("wmon")%></td>
</tr>
<%
	rs.movenext
Next
Set mypage=new xdownpage '创建对象 
 mypage.getconn=conn '得到数据库连接 
 mypage.getsql=sql
 mypage.pagesize=rs_sys("pagerecord") '设置每一页的记录条数据为5条 
 set rs=mypage.getrs() '返回Recordset 
 
 Response.Write "<table align=""center"" width='762px' border=0>"
 Response.Write "<tr>"
 Response.Write "<th colspan=6 align=""left"">"
 mypage.showpage()  '显示分页信息，这个方法可以，在set rs=mypage.getrs()以后任意位置调用，可以调用多次 
 'Response.Write "<label class=""button1""><input type=""submit"" value=""Excel""></label>"
 Response.Write "</th></tr></table>"
 Response.Write "</form>"
close_rs(rs_sys)
close_rs(rs)
endconnection
%>
</table>
</body>
</html>

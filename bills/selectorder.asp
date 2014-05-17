<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<script>
function choose(str1,str2){
	openwin('addbill.asp?type='+str1+'&billcode='+str2+'&chooseplan=true',900,600);	
	window.opener.top.close();
	window.close();	
}
</script>
<title>选择订单</title>
</head>

<body>
<%
s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
s_date2 = formatdate(date)
%>
<form name="form1" method="post" action="selectorder.asp?type=<%=request.QueryString("type")%>">
<table align="center">
<tr>
      <td width="50">从日期：</td>
      <td width="150"><input type="text" name="date1" size="16" value=<%If Request("date1") = "" Then
                                                              Response.Write Year(Date()) & "-" & Month(Date()) & "-1"
                                                            Else
                                                              Response.Write Request("date1")
                                                            End If%>><%showdate("date1")%></label></td>
      <td width="50">到日期：</td>
      <td width="150"><input type="text" name="date2" size="16" value=<%If Request("date2") = "" Then
                                                              Response.Write Date()
                                                            Else
                                                              Response.Write Request("date2")
                                                            End If%>><%showdate("date2")%></td>
      <td rowspan="2"><input type="submit" value=" 查 找 "  class="button"></td>
</tr>
<tr>
      <td width="50"><%if request.QueryString("type")="CG" then
	                     Response.Write("供应商：")
					   else
					     Response.Write("客 户：")
					   end if%></td>
      <td colspan="3"><input type="text" name="custname" id="cust" size="45" value=<%=Request.Form("custname")%> ><a href="#SelectDate" onClick="JavaScript:window.open ('../common/selectcust.asp', 'selectcust', 'height=600, width=800,top=100,left=150, toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=no')"><img border="0" src="../img/choose.gif" width="21" height="17"></a></td>
</tr>
</table>
</form>
<%
if Request.Form("date1") <> "" then
  s_date1 =  Request.Form("date1")
else
  s_date1 =  s_date1 
end if
if Request.Form("date2") <> "" then
  s_date2 =  Request.Form("date2") 
else
  s_date2 = s_date2 
end if
if Request.Form("custname") <> "" then
  s_cust = Request.Form("custname") 
else
  s_cust = ""
end if
if Request.QueryString("type") = "CG" then
  s_type = "采购订货"
else
  s_type = "销售订货"
end if
'sql = "select distinct '<input type=button value=选择此订单 class=button onClick=choose("""&request.QueryString("type")&""",'''+billcode,''')>' as action,s1.billcode,custname,depotname,memo,username,adddate from (select billcode+goodscode as bg,goodscode,sum(number) as t_num,billcode,custname,depotname,memo,username,adddate from s_billdetail where billtype = '"&s_type&"' and t_billdetail.check=1"& s_date1 & s_date2 & s_cust & " group by billcode,goodscode,custname,depotname,memo,username,adddate) as s1 left join (select planbillcode+goodscode as pg,planbillcode,goodscode,sum(number) as t_num from s_billdetail group by planbillcode,goodscode) as s2 on s1.bg = s2.pg where s1.t_num > s2.t_num or s2.t_num is null"

'sql="UP_SelectOrder @type='"&request.QueryString("type")&"',@s_type='"&s_type&"',@begindate='"&s_date1&"',@enddate ='"&s_date2&"',@cust='"&s_cust&"'"
sql = "select CONCAT('<input type=button value=选择此订单 class=button onClick=""choose(\'"
sql = sql & request.QueryString("type") & "\',\'',billcode,'\');"" />') as action,"
sql = sql & "(select sum(number) from t_billdetail where billcode = bill.billcode) as orderqty,"
sql = sql & "(select sum(number) from t_billdetail where billcode in (select billcode from t_bill where planbillcode = bill.billcode)) as finishqty,"
sql = sql & "billcode,custname,depotname,memo,username,adddate from t_bill as bill"
sql = sql & " where billtype = '"& s_type &"'"
sql = sql & " and bill.check = 1"
sql = sql & " and ifnull((select sum(number) from t_billdetail where billcode = bill.billcode),0) > ifnull((select sum(number) from t_billdetail where billcode in (select billcode from t_bill where planbillcode = bill.billcode)),0)"
call showpage(sql,"chooseorder",1)
%>
</body>
</html>

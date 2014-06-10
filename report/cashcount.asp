<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<%
if Request.QueryString("type") = "fk" then
call CheckAuthority("strCustFKCount")
end if
if Request.QueryString("type") = "sk" then
call CheckAuthority("strCustSKCount")
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<link rel="stylesheet" href="dtree.css" type="text/css">
<script type="text/javascript" src="../js/dtree.js"></script>
<script type="text/javascript" language="javascript">
function showProc(){
  message_box.style.visibility='visible';
  //创建灰色背景层
  procbg = document.createElement("div"); 
  procbg.setAttribute("id","mybg"); 
  procbg.style.background = "#f7f7f7"; 
  procbg.style.width = "100%"; 
  procbg.style.height = "100%"; 
  procbg.style.position = "absolute"; 
  procbg.style.top = "0"; 
  procbg.style.left = "0"; 
  procbg.style.zIndex = "500"; 
  procbg.style.opacity = "0.3"; 
  procbg.style.filter = "Alpha(opacity=30)"; 
  //背景层加入页面
  document.body.appendChild(procbg);
  document.body.style.overflow = "hidden";
}
//拖动
function drag(obj){  
     var s = obj.style;  
     var b = document.body;   
  var x = event.clientX + b.scrollLeft - s.pixelLeft;   
  var y = event.clientY + b.scrollTop - s.pixelTop; 
 
  var m = function(){  
   if(event.button == 1){  
    s.pixelLeft = event.clientX + b.scrollLeft - x;   
    s.pixelTop = event.clientY + b.scrollTop - y;   
   }else {
    document.detachEvent("onmousemove", m);
   }  
  }  
  document.attachEvent("onmousemove", m);
  if(!this.z) 
   this.z = 999;   
  s.zIndex = ++this.z;   
  event.cancelBubble = true;   
}
 
function closeProc(){
  message_box.style.visibility= "hidden";
  procbg.style.visibility = "hidden";
}
</script>
<title></title>
</head>

<body style="background:#FFFFFF">
<%
s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
s_date2 = formatdate(date)
%>
<div style="background:url(../images/main-02.gif); width:100%; height:20; padding-left:0" ><img src="../images/main-01.gif"></div>
<div style="padding-left:10px">
<br>
<%if Request.QueryString("type") = "FK" then%>
<br><div align="left"><span class="STYLE1">应付账款汇总表</span></div><br>
<%else%>
<br><div align="left"><span class="STYLE1">应收账款汇总表</span></div><br>
<%end if%>
<form id="form1" name="sample" method="post" action="cashcount.asp?type=<%=request.QueryString("type")%>">

  <table width="100%" border="0" id="tbl_tittle">
    <tr>
      <td width="70px" height="26" align="right">从&nbsp;日&nbsp;期：</td>
      <td width="150px"><label>
      <input type="text" name="date1" size="16" value=<%
If Request("date1") = "" Then
    Response.Write Year(Date()) & "-" & Month(Date()) & "-1"
Else
    Response.Write Request("date1")
End If

%>><%showdate("date1")%></label></td>
      <td width="100px" align="right">到&nbsp;日&nbsp;期：</td>
      <td width="200px"><input type="text" name="date2" size="16" value=<%
If Request("date2") = "" Then
    Response.Write Date()
Else
    Response.Write Request("date2")
End If

%>><%showdate("date2")%></td>
	<td rowspan="2"><input type="submit" value=" 查 询 " onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  class="button" /></td>
    </tr>
	<tr>
      <td align="right">往来单位：</td>
      <td><label>
      <input type="text" name="custname" id="cust" size="16" value="<%=Request.Form("custname")%>"><a href="#SelectDate" onClick="JavaScript:window.open ('../common/selectcust.asp', 'newwindow', 'height=600, width=800,top=100,left=150, toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=no')"><img border="0" src="../img/choose.gif" width="21" height="17"></a></label></td>
	<td align="right">往来单位类别：</td>
	<td><input type="hidden" name="typecode" id="typecode" value="<%=request.Form("typecode")%>">
      <input type="text" name="nodename" id="nodename" size="16" value=<%=Request.Form("nodename")%> ><a href="#" onClick="window.open ('../common/tree.asp?type=cust', 'newwindow', 'height=400, width=200,left=600,top=200,toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no, status=yes')"><img border="0" src="../img/choose.gif" width="21" height="17"></a>
	</td>
</tr>
</table>
<hr>
</form>
<%
If Request.Form("date1")<>"" Then
    s_date1 = Request.Form("date1")
End If
If Request.Form("date2")<>"" Then
    s_date2 = Request.Form("date2")
End If
if Request.QueryString("type") = "fk" then
sql = "select cust.custname, ifnull(cashbegin,0) - startmoneyf + startmoneys as cashbegin, ifnull(cashnow, 0) as cashnow, ifnull(paiednow,0) as paiednow, "
sql = sql + "ifnull(cashbegin,0) - startmoneyf + startmoneys + ifnull(cashnow, 0) - ifnull(paiednow,0) as cashend from "
sql = sql + "(select custname, startmoneyf, startmoneys from t_custom where custname like '"& Request.Form("custname") &"%' and code like '"& Request.Form("typecode") &"%') as cust "
sql = sql + "left join (select sum(wsmoney) as cashbegin, custname from s_pay where custname like '"& Request.Form("custname") &"%' and adddate <'"& s_date1 &"' group by custname) as start on start.custname = cust.custname "
sql = sql + "left join (select sum(wsmoney) as cashnow, custname from s_pay where custname like '"& Request.Form("custname") &"%' and adddate between '"& s_date1 &"' and '"& s_date2 &"' group by custname) as cash on cash.custname = cust.custname "
sql = sql + "left join (select sum(money * sign) as paiednow,custname from t_cash where custname like '"& Request.Form("custname") &"%' and adddate between '"& s_date1 &"' and '"& s_date2 &"' group by custname) as pay on pay.custname = cust.custname"
else

end if
if Request.QueryString("type") = "fk" then
	call showpage(sql,"cashcount_fk",1)
else
	call showpage(sql,"cashcount_sk",1)
end if
endconnection
%>
</div>

</body>
</html>

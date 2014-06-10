<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<% call CheckAuthority("strTotalAccount") %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<title>帐户余额查询</title>
</head>

<body>
<%
if Request.Form("date1") = "" then
  s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
else
  s_date1 = Request.Form("date1")
end if

if Request.Form("date2") = "" then
  s_date2 = formatdate(date)
else
  s_date2 = Request.Form("date2")
end if
%>
<br>
    <div align="left" class="STYLE1">帐户余额查询</div>
<table align="left" width="650" border="0">
<tr>
  <td>
    <form name="form1" action="Totalaccount.asp" method="post">
	<table align="left" width="500">
	<tr>
	  <td align="left" width="10%">帐&nbsp;&nbsp;&nbsp;&nbsp;户:</td>
      <td align="left" width="20%"><select name="account"><option value=""></option>
               <%
	           sql_account = "select * from t_account"
	    	       Set rs_account = server.CreateObject("adodb.recordset")
				  rs_account.open sql_account,conn,1,1
				  Do While rs_account.EOF = False
				   if request.Form("account")=rs_account("account") then
	 			     Response.Write "<option value="&rs_account("account")&" selected>"&rs_account("account")&"</option>"
				   else
				     Response.Write "<option value="&rs_account("account")&">"&rs_account("account")&"</option>"
				   end if
				   rs_account.movenext
				  loop
				  close_rs(rs_account)
				  %></select>
	  </td>
	  <td><input type="submit" value=" 查 询 " onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  class="button"></td>	  
	</tr>
	</table>
	</form>
  </td>
</tr>
<tr>
  <td>
  	<hr>
	<%
	if Request.Form("account") = "" then
	  s_account = ""
	else
	  s_account = " and account = '" & Request.Form("account") & "'"
	end if
	
sql = "select s1.account,original+ifnull(s2.t_mon,0)+ifnull(s3.t_mon,0)+ifnull(s4.t_mon,0) as balance from (select account,original from t_account where 1=1 " & s_account & ") as s1 left join (select account,sum(pay*flag*-1) as t_mon from t_bill where 1=1 " & s_account & " group by account) as s2 on s2.account=s1.account left join (select account,sum(money*sign) as t_mon from t_cash where 1=1 " & s_account & " group by account) as s3 on s3.account=s1.account left join (select account,sum(money*sign) as t_mon from t_cashbank where 1=1 " & s_account & " group by account) as s4 on s4.account = s1.account"
'response.write(sql)
Call showpage(sql,"Other1",1)
endconnection
%>
  </td>
</tr>
</table>
</body>
</html>

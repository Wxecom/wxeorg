<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<title>增加/修改帐户信息</title>
</head>

<body>
<%
if Request.QueryString("add") = "true" then
  s_tittle = "新增帐户信息"
  s_url    = "../action/add_account.asp?add=true"
  s_account = ""
  s_bank    = ""
  s_accountNO = ""
  d_original = 0
  s_memo    = ""
  s_id      = ""
else
  s_tittle = "修改帐户信息"
  s_url    = "../action/add_account.asp?add=false"
  sql = "select t_account.* from t_account where id = " & Request.QueryString("id")
  Set rs = server.CreateObject("adodb.recordset")
  rs.Open sql, conn, 1, 1
  s_account = rs("account")
  s_bank    = rs("bank")
  s_accountNO = rs("accountNO")
  d_original = rs("original")
  s_memo    = rs("memo")
  s_id      = rs("id")
end if
endconnection
%>
<br>
<div class="STYLE1" align="center"><%=s_tittle%></div>
<br>
<form name="form1" method="post" action="<%=s_url%>">
<input type="hidden" name="id" value="<%=s_id%>">
<table id="tbl" width="500" align="center">
  <tr>
    <td align="right" width="70">帐户名称:</td>
	<td><input type="text" name="account" size="16" value="<%=s_account%>"></td>
    <td align="right" width="70">开户银行:</td>
	<td><input type="text" name="bank" size="16" value="<%=s_bank%>"></td>
  </tr>
  <tr>
    <td align="right" width="70">帐&nbsp;&nbsp;&nbsp;&nbsp;号:</td>
	<td><input type="text" name="accountNO" size="16" value="<%=s_accountNO%>"></td>
    <td align="right" width="70">初始金额:</td>
	<td><input type="text" name="original" size="16" value="<%=d_original%>"></td>
  </tr>
  <tr>
    <td align="right" width="70">备&nbsp;&nbsp;&nbsp;&nbsp;注:</td>
	<td colspan="3"><input type="text" name="memo" size="50" value="<%=s_memo%>"></td>
  </tr>
</table>
<br>
<div align="center"><input class="button" type="submit"  onMouseOver="this.className='button_over';" onMouseOut="this.className='button';" value="保&nbsp;&nbsp;存">&nbsp;&nbsp;<input class="button"  onMouseOver="this.className='button_over';" onMouseOut="this.className='button';" type="reset" value="重&nbsp;&nbsp;置"></div>

</form>
</body>
</html>
<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../inc/config.asp" -->
<%call CheckAuthority("strReportProduct")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<title>产品入库报表</title>
</head>

<body>
<%
s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
s_date2 = formatdate(date)
%>
<form id="form1" name="sample" method="post" action="reportproduct.asp">
  <table width="100%" border="0" id="tbl_tittle">
    <tr>
      <td width="100" height="24" align="right">产品入库仓：</td>
      <td width="150">      <%
      Response.Write "<select name=""depot2"">"
	  Response.Write "<option value=""""></option>"
      sql_depot = "select depotname from t_user where username='"&request.cookies("username")&"'"
      Set rs_depot = conn.Execute(sql_depot)
      arr = split(rs_depot("depotname"),",")
      if ubound(arr) <> -1 then
        for i = lbound(arr) to ubound(arr)-1
          s_depotpower = s_depotpower & "'" & arr(i) & "',"
	    next
      end if
      sql_depot = "select * from t_depot where depotname in (" & s_depotpower & "'')"
      set rs_s_depot = conn.execute(sql_depot)
      Do While rs_s_depot.EOF = False
        if Request.Form("depot2") = rs_s_depot("depotname") then
	      Response.Write "<option value="&rs_s_depot("depotname")&" selected>"&rs_s_depot("depotname")&"</option>"
	    else
	      Response.Write "<option value="&rs_s_depot("depotname")&">"&rs_s_depot("depotname")&"</option>"
	    end if
	   rs_s_depot.movenext
      loop
      Response.Write "</select>"
	  close_rs(rs_depot)
	  close_rs(rs_s_depot)
	  %></td>
      <td width="100" align="right">原材料出库仓：</td>
      <td width="150">      <%
      Response.Write "<select name=""depot1"">"
	  Response.Write "<option value=""""></option>"
      sql_depot = "select depotname from t_user where username='"&request.cookies("username")&"'"
      Set rs_depot = conn.Execute(sql_depot)
      arr = split(rs_depot("depotname"),",")
      if ubound(arr) <> -1 then
        for i = lbound(arr) to ubound(arr)-1
          s_depotpower = s_depotpower & "'" & arr(i) & "',"
	    next
      end if
      sql_depot = "select * from t_depot where depotname in (" & s_depotpower & "'')"
      set rs_s_depot = conn.execute(sql_depot)
      Do While rs_s_depot.EOF = False
        if Request.Form("depot1") = rs_s_depot("depotname") then
	      Response.Write "<option value="&rs_s_depot("depotname")&" selected>"&rs_s_depot("depotname")&"</option>"
	    else
	      Response.Write "<option value="&rs_s_depot("depotname")&">"&rs_s_depot("depotname")&"</option>"
	    end if
	   rs_s_depot.movenext
      loop
      Response.Write "</select>"
	  close_rs(rs_depot)
	  close_rs(rs_s_depot)
	  %></td>
      <td rowspan="3"><input type="submit" name="Submit" value=" 查 询 " class="button" /></td>
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
	<tr><td align="right">产品名称：</td><td><input type="text" name="goodsname" id="goodsname" size="16"></td>
	<td align="right">经&nbsp;办&nbsp;人：</td><td><select name="user" id="user">
	<option value=""></option>
    <%
sql = "select * from t_Employee"
Set a = conn.Execute(sql)
Do While a.EOF = False
    s_name = a("name")
    If Request.Form("user") = a("name") Then
        Response.Write "<option value="&s_name&" selected>"&s_name&"</option>"
    Else
        Response.Write "<option value="&s_name&">"&s_name&"</option>"
    End If
    a.movenext
Loop
close_rs(a)
%>
	</select></td></tr>
  </table>
</form>
<div align="center" class="style1">产品入库报表</div>
<%
If Request.Form("date1")<>"" Then
    s_date1 = Request.Form("date1")
End If
If Request.Form("date2")<>"" Then
    s_date2 = Request.Form("date2")
End If
If Request.Form("depot1") = "" Then
    s_depotname = " and custname in ("&s_depotpower&"'')"
Else
    s_depotname = " and custname='"&Request.Form("depot1")&"'"
End If
If Request.Form("depot2") = "" Then
    s_custname = " and depotname in ("&s_depotpower&"'')"
Else
    s_custname = " and depotname='"&Request.Form("depot2")&"'"
End If
If Request.Form("goodsname") = "" Then
    s_goodsname = ""
Else
    s_goodsname = " and goodsname like '%"&Request.Form("goodsname")&"%'"
End If
If Request.Form("user") = "" Then
    s_user = ""
Else
    s_user = " and username='"&Request.Form("user")&"'"
End If
sql = "select number * flag as t_num,s_billdetail.* from s_billdetail where flag=1 and billtype = '产品组装' and adddate>='"&s_date1&"' and adddate<='"&s_date2&"'" & s_depotname & s_custname & goodsname & s_user
call showpage(sql,"R_Detail",1)
endconnection
%>
</body>
</html>

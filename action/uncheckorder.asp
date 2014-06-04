<!-- #include file="../inc/conn.asp" -->
<!-- #include file="checkuser.asp" -->
<%
if Request.QueryString("type") = "CG" then
  call CheckAuthority("strUnCheckCD")
end if

if  Request.QueryString("type") = "XS" then
  call CheckAuthority("strUnCheckXD")
end if

s_billcode = Request.QueryString("billcode")
sql = "select * from t_bill where billcode = '" & s_billcode & "'"
Set rs = conn.Execute(sql)
if rs("check") = 0 then%>
<script language=javascript>
alert("该单据已反审!");
window.history.go(-1);
//window.reload();
</script>
<%else
sql = "update t_bill set t_bill.check=0 where billcode = '" & s_billcode & "'"
Set rs = conn.Execute(sql)%>
<script language=javascript>
alert("该单据已反审!");
window.history.go(-1);
//window.reload();
</script>
<%end if
endconnection%>
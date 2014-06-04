<!-- #include file="../inc/conn.asp" -->
<!-- #include file="checkuser.asp" -->
<%
if Request.QueryString("type") = "CD" then
  call CheckAuthority("strCheckCD")
end if

if  Request.QueryString("type") = "XD" then
  call CheckAuthority("strCheckXD")
end if

s_billcode = Request.QueryString("billcode")
sql = "select * from t_bill where billcode = '" & s_billcode & "'"
Set rs = conn.Execute(sql)
if rs("check") = 1 then%>
<script language=javascript>
alert("该单据已审核！");
window.history.go(-1)
</script>
<%else
sql = "update t_bill set t_bill.check=1 where billcode = '" & s_billcode & "'"
Set rs = conn.Execute(sql)%>
<script language=javascript>
alert("该单据已审核！");
window.history.go(-1)
</script>
<%end if
endconnection%>
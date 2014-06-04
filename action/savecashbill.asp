<!-- #include file="../inc/conn.asp" -->
<!-- #include file="checkuser.asp" -->
<%
if Request.QueryString("type") = "FK" then
    i_sign = -1
    s_cashtype = "付款"
else
    i_sign = 1
    s_cashtype = "收款"
end if

set rsCash = Server.CreateObject("adodb.recordset")
sql = "select * from t_cash where cashcode = '"& Request.Form("billcode") &"'"
rsCash.Open sql, conn, 1, 3
if Request.QueryString("add") then
	rsCash.addnew
end if
rsCash("cashcode") = Request.Form("billcode")
rsCash("billcode") = Request.Form("bill")
rsCash("type") = Request.Form("type")
rsCash("money") = Request.Form("money")
rsCash("memo") = Request.Form("remark")
rsCash("user") = Request.Form("user")
rsCash("depart") = Request.Form("depart")
rsCash("custname") = Request.Form("Cust")
rsCash("cashtype") = s_cashtype
rsCash("adddate") = Request.Form("date")
rsCash("account") = Request.Form("account")
rsCash("sign") = i_sign
rsCash("username") = Request.Form("maker")
rsCash("debit") = Request.Form("debit")
rsCash("lender") = Request.Form("lender")
rsCash.update
rsCash.close
set rsCash = nothing
endconnection
%>
<script language=javascript>
alert('保存成功！');
window.close();
</script>


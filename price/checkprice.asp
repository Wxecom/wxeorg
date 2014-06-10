<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<%
on error resume next
conn.BeginTrans
sBillCode = Request.QueryString("billcode")
sql = "update t_price set t_price.check = 1 where billcode = '"& sBillCode &"'"
conn.Execute(sql)

sqlprice = "select * from t_price where billcode = '"& sBillCode &"'"
set rsprice = Server.CreateObject("adodb.recordset")
rsprice.open sqlprice, conn, 1, 1
sqldepot = "select * from t_stock where goodscode = '"& rsprice("goodscode") &"' and depotname = '"& rsprice("depotname") &"'"
set rsdepot = Server.CreateObject("adodb.recordset")
rsdepot.open sqldepot, conn, 1, 3
if rsdepot.recordcount = 0 then
	rsdepot.addnew
	rsdepot("goodscode") = rsprice("goodscode")
	rsdepot("depotname") = rsprice("depotname")
	rsdepot("qty") = rsprice("newqty")
	rsdepot("price") = rsprice("newprice")
	rsdepot.update
else
	rsdepot("goodscode") = rsprice("goodscode")
	rsdepot("depotname") = rsprice("depotname")
	rsdepot("qty") = rsprice("newqty")
	rsdepot("price") = rsprice("newprice")
	rsdepot.update
end if
if err <> 0 then
	Response.Write("<script>alert('审核时出现错误!');</script>")
	conn.rollbacktrans
else
	Response.Write("<script>alert('该单据审核成功！');</script>")
	conn.CommitTrans
end if
endconnection
%>
<script>
window.opener.location.reload();
window.close();
</script>
<!-- #include file="../inc/conn.asp" -->
<!-- #include file="checkuser.asp" -->
<%
sql = "select billtype from t_bill where billcode='"&Request.QueryString("billcode")&"'"
Set rs = server.CreateObject("adodb.recordset")
rs.Open sql, conn, 1, 1
if rs("billtype") = "采购订货" then
	call CheckAuthority("strDelCD")
end if
if rs("billtype") = "销售订货" then
	call CheckAuthority(str"DelXD")
end if
if rs("billtype") = "采购入库" then
	call CheckAuthority("strDelCG")
end if
if rs("billtype") = "销售出库" then
	call CheckAuthority("strDelXS")
end if
if rs("billtype") = "采购退货" then
	call CheckAuthority("strDelCT")
end if
if rs("billtype") = "销售退货" then
	call CheckAuthority("strDelXT")
end if
if rs("billtype") = "其他入库" then
	call CheckAuthority("strDelIn")
end if
if rs("billtype") = "其他出库" then
	call CheckAuthority("strDelOut")
end if
if rs("billtype") = "仓库调拨" then
	call CheckAuthority("strDelDB")
end if
if rs("billtype") = "库存盘点" then
	call CheckAuthority("strDelPD")
end if
if rs("billtype") = "组装出库" then
	call CheckAuthority("strDelZZ")
end if
if rs("billtype") = "领料出库" then
	call CheckAuthority("strDelLL")
end if
if rs("billtype") = "退料入库" then
	call CheckAuthority("strDelTL")
end if

sql = "select t_bill.check from t_bill where billcode = '" & Request.QueryString("billcode") & "'"
Set rs = conn.Execute(sql)


if rs("check") = True then
  Response.Write "<script>alert(""审核过的单据不能删除，请先反审!"");</script>"
  Response.Write "<script>window.history.go(-1);</script>"
  Response.End()
end if
sql = "delete from t_bill where billcode='"&Request.QueryString("billcode")&"'"
q="delete from t_cash where billcode='"&Request.QueryString("billcode")&"'"
conn.Execute(sql)
conn.execute(q)
close_rs(rs)
endconnection
%>
<script>
	window.history.go(-1);
</script>

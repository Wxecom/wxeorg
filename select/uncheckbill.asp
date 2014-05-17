<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>   
<!-- #include file="../inc/conn.asp" -->
<%
Response.ContentType = "text/html"
Response.Charset = "gb2312"
s_billcode = Request("billcode")




'检查相关单据
sqlcheck = "select count(billcode) as c from t_bill where planbillcode = '"& s_billcode &"'"
set rscheck = Server.CreateObject("adodb.recordset")
rscheck.open sqlcheck, conn, 1, 1
if CInt(rscheck("c")) > 0 then
	Response.Write s_billcode & ":存在相关单据"
	Response.End()
end if
rscheck.close()
set rscheck = nothing
'检查收付款记录
sqlcheck = "select count(billcode) as c from t_cash where billcode = '"& s_billcode &"'"
set rscheck = Server.CreateObject("adodb.recordset")
rscheck.open sqlcheck, conn, 1, 1
if CInt(rscheck("c")) > 0 then
	Response.Write s_billcode & ":有收付款记录"
	Response.End()
end if
rscheck.close()
set rscheck = nothing
sqlbill = "select billcode,depotname,flag,custname,billtype,t_bill.check from t_bill where billcode = '"& s_billcode &"'"
set rsbill = server.CreateObject("adodb.recordset")
rsbill.open sqlbill, conn, 1, 1
if rsbill("billtype") = "采购订货" then
	power = Authority("strUnCheckCD")
end if
if rsbill("billtype") = "销售订货" then
	power = Authority("strUnCheckXD")
end if
if rsbill("billtype") = "采购入库" then
	power = Authority("strUnCheckCG")
end if
if rsbill("billtype") = "销售出库" then
	power = Authority("strUnCheckXS")
end if
if rsbill("billtype") = "采购退货" then
	power = Authority("strUnCheckCT")
end if
if rsbill("billtype") = "销售退货" then
	power = Authority("strUnCheckXT")
end if
if rsbill("billtype") = "其他入库" then
	power = Authority("strUnCheckIn")
end if
if rsbill("billtype") = "其他出库" then
	power = Authority("strUnCheckOut")
end if
if rsbill("billtype") = "仓库调拨" then
	power = Authority("strUnCheckDB")
end if
if rsbill("billtype") = "库存盘点" then
	power = Authority("strUnCheckPD")
end if
if rsbill("billtype") = "组装出库" then
	power = Authority("strUnCheckZZ")
end if
if rsbill("billtype") = "领料出库" then
	power = Authority("strUnCheckLL")
end if
if rsbill("billtype") = "退料入库" then
	power = Authority("strUnCheckTL")
end if
if power = "False" then
	Response.Write s_billcode & ":无反审权限"
	Response.End()
end if            

if rsbill("check") = "False" then   
	Response.Write s_billcode & ":请勿重复反审"
	Response.End()
end if
  

if rsbill("flag") <> "0" then
sqldetail = "select goodscode,number,price,inprice from t_billdetail where billcode = '"& s_billcode &"'"
Set rsdetail = server.CreateObject("adodb.recordset")
rsdetail.Open sqldetail, conn, 1,1
rsdetail.movefirst
sql = "select * from t_inout where id < 0"
Set rs_inout = server.CreateObject("adodb.recordset")
rs_inout.Open sql, conn, 1, 1
do while not rsdetail.eof
	rs_inout.addnew
	rs_inout("billcode") = rsbill("billcode")
	rs_inout("goodscode") = rsdetail("goodscode")
	rs_inout("qty") = rsdetail("number")
	rs_inout("price") = rsdetail("price")
	rs_inout("inprice") = rsdetail("inprice")
	rs_inout("depotname") = rsbill("depotname")
	rs_inout("flag") = -1 * rsbill("flag")
	rs_inout.update
	rsdetail.movenext
loop


if rsbill("BillType") = "仓库调拨" then
rsdetail.movefirst
do while not rsdetail.eof
	rs_inout.addnew
	rs_inout("billcode") = rsbill("billcode")
	rs_inout("goodscode") = rsdetail("goodscode")
	rs_inout("qty") = rsdetail("number")
	rs_inout("price") = rsdetail("price")
	rs_inout("inprice") = rsdetail("inprice")
	rs_inout("depotname") = rsbill("custname")
	rs_inout("flag") = rsbill("flag")
	rs_inout.update
	rsdetail.movenext
loop  
end if
end if
on error resume next
conn.BeginTrans   
sql = "update t_bill set t_bill.Check=0,checkman='"& Request.Cookies("username") &"' where billcode = '" & s_billcode & "'"
Set rs = conn.Execute(sql)  
if err <> 0 then
	conn.rollbacktrans
	Response.Write "False"
else
	conn.CommitTrans
	Response.Write "True"
end if
Response.End()
endconnection
%>
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>   
<!-- #include file="../inc/conn.asp" -->
<%
Response.ContentType = "text/html"
Response.Charset = "gb2312"

on error resume next
conn.BeginTrans
s_billcode = Request("billcode")
sqlbill = "select billcode,depotname,flag,custname,billtype,t_bill.check from t_bill where billcode = '"& s_billcode &"'"
set rsbill = server.CreateObject("adodb.recordset")
rsbill.open sqlbill, conn, 1, 1
s_billtype=rsbill("billtype")
if s_billtype = "采购订货" then
	power = Authority("strCheckCD")
end if
if s_billtype = "销售订货" then
	power = Authority("strCheckXD")
end if
if s_billtype = "采购入库" then
	power = Authority("strCheckCG")
end if
if s_billtype = "销售出库" then
	power = Authority("strCheckXS")
end if
if s_billtype = "采购退货" then
	power = Authority("strCheckCT")
end if
if s_billtype = "销售退货" then
	power = Authority("strCheckXT")
end if
if s_billtype = "其他入库" then
	power = Authority("strCheckIn")
end if
if s_billtype = "其他出库" then
	power = Authority("strCheckOut")
end if
if s_billtype = "仓库调拨" then
	power = Authority("strCheckDB")
end if
if s_billtype = "库存盘点" then
	power = Authority("strCheckPD")
end if
if s_billtype = "组装出库" then
	power = Authority("strCheckZZ")
end if
if s_billtype = "领料出库" then
	power = Authority("strCheckLL")
end if
if s_billtype = "退料入库" then
	power = Authority("strCheckTL")
end if

if power = "False" then
	Response.Write s_billcode & ":无审核权限"
	Response.End()
end if



if rsbill("check") = "True" then
	Response.Write s_billcode & ":请勿重复审核"
	Response.End()
end if


if rsbill("flag") <> "0" then
sqldetail = "select goodscode,number,price,inprice from t_billdetail where billcode = '"& s_billcode &"'"


Set rsdetail = server.CreateObject("adodb.recordset")
rsdetail.Open sqldetail, conn, 1,3

rsdetail.movefirst
sql = "select * from t_inout where id < 0"
Set rs_inout = server.CreateObject("adodb.recordset")
rs_inout.Open sql, conn, 1, 3

Do While not rsdetail.eof=false

	rs_inout.addnew
	rs_inout("billcode") = rsbill("billcode")
	rs_inout("goodscode") = rsdetail("goodscode")
	rs_inout("qty") = rsdetail("number")
	rs_inout("price") = rsdetail("price")
	if rsbill("flag") = "-1" then
		'审核时取最新成本价
		sql = "select price,qty from t_stock where goodscode = '"& rsdetail("goodscode") &"' and depotname = '"& rsbill("depotname") &"'"
		set rs_stock = server.CreateObject("adodb.recordset")
		rs_stock.open sql, conn, 1, 3
		
		
		
		
		'判断负出库
		if (CInt(rsbill("flag")) * CInt(rsdetail("number"))) < 0 then
			if cdbl(rsdetail("number")) > cdbl(rs_stock("qty")) then
				Response.Write s_billcode & ":有负出库的情况("& rsdetail("goodscode") &"实际库存数量是"&rs_stock("qty")&",您要求的出库数量为"&rsdetail("number")&")"
				Response.End()
			end if
		end if
		
		if rs_stock.recordcount = 0 then
			rs_inout("inprice") = 0
			rsdetail("inprice") = 0
		else
			rs_inout("inprice") = rs_stock("price")
			rsdetail("inprice") = rs_stock("price")
		end if
		
		rsdetail.update
		
	else
		rs_inout("inprice") = rsdetail("price")
	end if
	
	rs_inout("depotname") = rsbill("depotname")
	rs_inout("flag") = rsbill("flag")
	rs_inout.update
	rsdetail.movenext
	
loop

'if s_billtype = "仓库调拨" then
'Set rsdetails = server.CreateObject("adodb.recordset")
'rsdetails.Open sqldetail, conn, 1,3
'rsdetails.movefirst '这个代码运行不了，可能有问题，需要调试
'while not rsdetails.eof
'	rs_inout.addnew
'	rs_inout("billcode") = rsbill("billcode")
'	rs_inout("goodscode") = rsdetails("goodscode")
'	rs_inout("qty") = rsdetails("number")
'	rs_inout("price") = rsdetails("price")
'	rs_inout("inprice") = rsdetails("inprice")
'	rs_inout("depotname") = rsbill("custname")
'	rs_inout("flag") = -1*rsbill("flag")
'	rs_inout.update
'	rsdetail.movenext
'loop
'
'end if
'
end if


sql = "update t_bill set t_bill.check=1,checkman='"& Request.Cookies("username") &"',checkdate='"& date() &"' where billcode = '" & s_billcode & "'"
Set rs = conn.Execute(sql)
if err <> 0 then
	Response.Write "False"
	conn.rollbacktrans
else
	Response.Write "True"
	conn.CommitTrans
end if

Response.End()
endconnection
%>
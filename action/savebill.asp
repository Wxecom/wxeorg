<!-- #include file="../inc/conn.asp" -->
<!-- #include file="checkuser.asp" -->
<%
If Request("add") = "false" Then
	s_billtype = Request.QueryString("type")
	s_adddate = Trim(Request.Form("date"))
	s_custname = Trim(Request.Form("cust"))
	s_depot = Trim(Request.Form("depot"))
	s_maker = trim(Request.Form("maker"))
	s_user = Trim(Request.Form("user"))
	s_memo = Trim(Request.Form("memo"))
	s_pay = Trim(Request.Form("pay"))
	s_account = Trim(Request.Form("account"))
	s_billcode = Trim(Request.Form("billcode"))
	i_rowcount = Trim(Request.Form("rowcount"))
	s_zkprice=trim(request.Form("zkprice"))
	s_yfprice=trim(request.Form("yfprice"))
	s_zdprice=trim(request.Form("zdprice"))
	if s_pay = "" then s_pay = "0" end if
	if s_zkprice = "" then s_zkprice="0" end if
	if s_yfprice = "" then s_yfprice="0" end if
	if s_zdprice = "" then s_zdprice="0" end if
	if s_account = "" then s_account = "" end if
	if s_pay = "" then s_pay = "0" end if
	'添加开始
	on error resume next
	conn.BeginTrans
	set rs = Server.CreateObject("adodb.recordset")
	sql = "select * from dict_bill where name = '"& s_billtype &"'"
	rs.open sql, conn, 1, 1
	
	sqlup = "update t_bill set adddate = '" &s_adddate& "' ,adddate='" &s_adddate& "',custname= '" &s_custname& "',depotname='" &s_depot& "',username= '" &s_user& "',memo='" &s_memo& "',account='" &s_account& "',zdprice='" &s_zdprice& "',zkprice='" &s_zkprice& "',yfprice='" &s_yfprice& "',pay='" &s_pay& "'  where billcode = '" & s_billcode & "'"
	'response.write sqlup
	conn.Execute(sqlup)
	'conn.BeginTrans
	
Else
	set rs = Server.CreateObject("adodb.recordset")
	sql = "select * from dict_bill where billtype = '"& Request.QueryString("type") &"'"
	rs.open sql, conn, 1, 1
	
	today = Now()
	tdate = Year(today) & "-" & Right("0" & Month(today), 2) & "-" & Right("0" & Day(today), 2)
	set rsCode = Server.CreateObject("adodb.recordset")
	sql = "select * from t_bill where billcode like '"&rs("billtype")&"-"&tdate&"%' order by billcode desc"
	rsCode.Open sql, conn, 1, 1
	If rsCode.recordcount = 0 Then
		s_billcode = rs("billtype")&"-"&tdate&"-0001"
	Else
		s_temp = rsCode("billcode")
		s_temp = Right(s_temp, 4) + 1
		s_billcode = rs("billtype")&"-"&tdate&"-"&Right("000"&s_temp, 4)
	End If
	
	s_billtype = rs("name")
    s_adddate = Trim(Request.Form("date"))
    s_custname = Trim(Request.Form("cust"))
    s_depot = Trim(Request.Form("depot"))
	s_maker = trim(Request.Form("maker"))
    s_user = Trim(Request.Form("user"))
    s_memo = Trim(Request.Form("memo"))
	s_pay = Trim(Request.Form("pay"))
	s_account = Trim(Request.Form("account"))
    's_billcode = Trim(Request.Form("billcode"))
    i_rowcount = Trim(Request.Form("rowcount"))
	s_planbillcode = trim(Request.Form("planbillcode"))
	s_zkprice=trim(request.Form("zkprice"))
	s_yfprice=trim(request.Form("yfprice"))
	s_zdprice=trim(request.Form("zdprice"))
	if s_zkprice = "" then s_zkprice="0" end if
	if s_yfprice = "" then s_yfprice="0" end if
	if s_zdprice = "" then s_zdprice="0" end if
	if s_account = "" then s_account = "" end if
	if s_pay = "" then s_pay = "0" end if
	'添加开始
	on error resume next
	conn.BeginTrans
	
	sqladd="insert into t_bill(billtype,billcode,planbillcode,adddate,custname,depotname,username,memo,flag,account,zdprice,zkprice,yfprice,pay)VALUES('" &s_billtype & "','" &s_billcode & "','" & s_planbillcode & "','" & s_adddate & "','" & s_custname & "','" & s_depot & "','" & s_user& "','" & s_memo & "','" & rs("billflag") & "','" & s_account & "','" & s_zdprice & "','" & s_zkprice & "','" & s_yfprice & "','" & s_pay & "')"
	conn.Execute(sqladd)
	
End If
	
sql = "delete from t_billdetail where billcode='"&s_billcode&"';"
conn.Execute(sql)
set rsMemory = Server.CreateObject("adodb.recordset")

'set rsDetail = Server.CreateObject("adodb.recordset")
'sql = "select * from t_billdetail where billcode = '" & s_billcode & "'"
'rsDetail.Open sql, conn, 1, 3
arrGoodscode = split(Trim(Request.Form("goodscode")), ",")
arrGoodsname = split(Trim(Request.Form("goodsname")), ",")
arrGoodsunit = split(Trim(Request.Form("goodsunit")), ",")
arrUnits     = split(Trim(Request.Form("units")), ",")
arrPrice     = split(Trim(Request.Form("price")), ",")
arrNumber    = split(Trim(Request.Form("number")), ",")
arrMoney     = split(Trim(Request.Form("money")), ",")
arrRemark    = split(Trim(Request.Form("remark")), ",")
arrAvgprice  = split(Trim(Request.Form("aveprice")), ",")

if UBound(arrGoodscode) = 0 then
	
	if rs("inorout") = "入库" then
		sinprice = cdbl(Trim(Request.Form("price")))
	else
		sinprice = cdbl(Trim(Request.Form("aveprice")))
	end if
		sqldetail="insert into t_billdetail(billcode,goodscode,goodsname,goodsunit,units,price,number,money,detailnote,inprice)"&_
	          "values('"&s_billcode&"','"&Trim(Request.Form("goodscode"))&"','"&Trim(Request.Form("goodsname"))&_
	          "','"&Trim(Request.Form("goodsunit"))&"','"&Trim(Request.Form("units"))&"','"&cdbl(Trim(Request.Form("price")))&_
	          "','"&cdbl(Trim(Request.Form("number")))&"','"&cdbl(Trim(Request.Form("money")))&"','"&Trim(Trim(Request.Form("remark")))&_
	          "','"&sinprice&"');"
	          
	
	
	sql = "select * from t_memoryprice where goodscode = '"& Request.Form("goodscode") &"' and billtype = '"& s_billtype &"' and custname = '"& s_custname &"'"
	rsMemory.open sql, conn, 1, 3
	if rsMemory.eof then
		
		sqlnew="insert into t_memoryprice(goodscode,custname,billtype,price) values"&_
		"('"&Trim(Request.Form("goodscode"))&"','"&s_custname&"','"&s_billtype&"','"&cdbl(Trim(Request.Form("price")))&"');"
	  conn.Execute(sqlnew)
	end if
	rsMemory.close
	conn.Execute(sqldetail)
	
else
	For i = LBound(arrGoodscode) To UBound(arrGoodscode)
		
		if rs("inorout") = "入库" then
			sinprice = cdbl(Trim(arrPrice(i)))
		else
			sinprice = cdbl(Trim(arrAvgprice(i)))
		end if
		sqldetail="insert into t_billdetail(billcode,goodscode,goodsname,goodsunit,units,price,number,money,detailnote,inprice)"&_
	          "values('"&s_billcode&"','"&Trim(arrGoodscode(i))&"','"&Trim(arrGoodsname(i))&_
	          "','"&Trim(arrGoodsunit(i))&"','"&Trim(arrUnits(i))&"','"&cdbl(Trim(arrPrice(i)))&_
	          "','"&cdbl(Trim(arrNumber(i)))&"','"&cdbl(Trim(arrMoney(i)))&"','"&Trim(arrRemark(i))&_
	          "','"&sinprice&"');"
		
		
	sql = "select * from t_memoryprice where goodscode = '"& Trim(arrGoodscode(i)) &"' and billtype = '"& s_billtype &"' and custname = '"& s_custname &"'"
	rsMemory.open sql, conn, 1, 3
	if rsMemory.eof then
		
		sqlnew="insert into t_memoryprice(goodscode,custname,billtype,price) values"&_
		"('"&Trim(arrGoodscode(i))&"','"&s_custname&"','"&s_billtype&"','"&cdbl(Trim(arrPrice(i)))&"');"
	 conn.Execute(sqlnew)
	end if
	
	rsMemory.close
	
	conn.Execute(sqldetail)
	
	Next
end if
set rsMemory = nothing

rs.close
set rs = nothing

if err.number <= 0 then
	conn.CommitTrans
	conn.close
	set conn=nothing
	if (Request("type") = "CG") or (Request("type")="XS") then
		response.write "<script>alert('保存成功！');location.href='../bills/addbill.asp?type="&request.QueryString("type")&"';</script>"
	elseif (Request("type") = "CD") or (Request("type")="XD") then
		response.write "<script>alert('保存成功！');location.href='../bills/orderbill.asp?type="&request.QueryString("type")&"';</script>"
	elseif request("type")="RK" or request("type")="CK" then
		response.write "<script>alert('保存成功！');location.href='../bills/kc_depotbill.asp?type="&request.QueryString("type")&"';</script>"
	elseif request("type")="DB" then
		response.write "<script>alert('保存成功！');location.href='../bills/kc_depotbill.asp?type="&request.QueryString("type")&"&bill=bill';</script>"
	elseif request("type")="PD" then
		response.write "<script>alert('保存成功！');location.href='../bills/kc_depotbill.asp?type="&request.QueryString("type")&"&bill=pbill';</script>"
	elseif request("type")="LL" then
		response.write "<script>alert('保存成功！');location.href='../bills/kc_depotbill.asp?type="&request.QueryString("type")&"&bill=lbill';</script>"
	elseif request("type")="TL" then
		response.write "<script>alert('保存成功！');location.href='../bills/kc_depotbill.asp?type="&request.QueryString("type")&"&bill=lbill';</script>"
	else
		response.write "<script>alert('保存成功！');window.close();</script>"
	end if
else
	conn.RollbackTrans '否则回滚
	conn.close
	set conn=nothing
	response.write err.description
	response.write "<script>alert('保存失败！');window.close();</script>"
end if
%>

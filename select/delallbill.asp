<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>   
<!-- #include file="../inc/conn.asp" -->
<%
Response.ContentType = "text/html"
Response.Charset = "gb2312"
billcode=unescape(request("billcode"))
set rsDel = Server.CreateObject("adodb.recordset")
sql = "select * from t_bill where billcode = '" & billcode & "'"
rsDel.open sql, conn, 1, 3
if rsDel("billtype") = "采购订货" then
	power = Authority("strDelCD")
end if
if rsDel("billtype") = "销售订货" then
	power = Authority("strDelXD")
end if
if rsDel("billtype") = "采购入库" then
	power = Authority("strDelCG")
end if
if rsDel("billtype") = "销售出库" then
	power = Authority("strDelXS")
end if
if rsDel("billtype") = "采购退货" then
	power = Authority("strDelCT")
end if
if rsDel("billtype") = "销售退货" then
	power = Authority("strDelXT")
end if
if rsDel("billtype") = "其他入库" then
	power = Authority("strDelIn")
end if
if rsDel("billtype") = "其他出库" then
	power = Authority("strDelOut")
end if
if rsDel("billtype") = "仓库调拨" then
	power = Authority("strDelDB")
end if
if rsDel("billtype") = "库存盘点" then
	power = Authority("strDelPD")
end if
if rsDel("billtype") = "组装出库" then
	power = Authority("strDelZZ")
end if
if rsDel("billtype") = "领料出库" then
	power = Authority("strDelLL")
end if
if rsDel("billtype") = "退料入库" then
	power = Authority("strDelTL")
end if
if power = "False" then
	Response.Write billcode & ":无删除权限"
else
	if rsDel("check") = false then
	on error resume next
    conn.BeginTrans   

		delsql = "delete from t_billdetail where billcode = '" & billcode & "'"
		Set rs1 = conn.Execute(delsql)
		
		delsql = "delete from t_bill where billcode = '" & billcode & "'"
		Set rs2 = conn.Execute(delsql)
		
		if err <> 0 then        
			conn.rollbacktrans    
			Response.Write "False"
		else                    
			conn.CommitTrans      
			Response.Write "True" 
		end if                  

		
	else
		Response.Write billcode & ":不能删除已审核单据"
	end if
end if
rsDel.close
set rsDel = nothing
Response.End
'Response.Write s_account
%>delete from t_bill where billcode = 'XS-2014-05-17-0002'
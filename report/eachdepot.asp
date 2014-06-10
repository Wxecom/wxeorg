<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<% call CheckAuthority("strEachDepot") %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<title>分仓库存</title>
</head>

<body style="padding:0px;margin:0px;">

<%
sql = "select t_depot.depotname,goods.goodscode,goodsname,goodsunit,units,ifnull(qty,0) as qty,ifnull(price,0) as price,ifnull(qty*price,0) as depotamt from ( "
sql = sql + "t_depot left join "
sql = sql + "(select * from t_stock where goodscode = '"& Request.QueryString("goodscode") &"') as depot on depot.depotname=t_depot.depotname "
sql = sql + "left join (select goodscode,goodsname,goodsunit,units from t_Goods where goodscode = '"& Request.QueryString("goodscode") &"') as goods on depot.goodscode = goods.goodscode)"

call showpage(sql,"rgoodsdepotdetail",1)

endconnection
%>
</body>
</html>
<!-- #include file="../inc/conn.asp" -->
<HTML><HEAD><TITLE></TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<style type="text/css">
* {margin:0px;padding:0px;}
html,body {height:100%;}
body {font-size:14px; line-height:24px;}
#tip {
	position: absolute;
	right: 0px;
	bottom: 0px;
	height: 0px;
	width: 240px;
	border: 1px solid #CCCCCC;
	padding: 1px;
	overflow:hidden;
	display:none;
	font-size:12px;
	z-index:10;
}
#tip p {padding:1px;}
#tip h1 {font-size:12px;height:25px;line-height:25px;background-color:#0066CC;color:#FFFFFF;padding:0px 3px 0px 3px;}
#tip h1 a {float:right;text-decoration:none;color:#FFFFFF;}
</style>
<script language="JavaScript" src="../js/jquery.js"></script>
<script language="JavaScript" src="../js/mainjs.js"></script>
<script language=javascript>



var handle;
function start()
{
var obj = document.getElementById("tip");
document.getElementById('tip').style.height='0px';
if (parseInt(obj.style.height)==0)
{ obj.style.display="block";
handle = setInterval("changeH('up')",2);
}
}
function end()
{
     handle = setInterval("changeH('down')",2);
} 
function changeH(str)
{
var obj = document.getElementById("tip");
if(str=="up")
{
if (parseInt(obj.style.height)>150)
   clearInterval(handle);
else
   obj.style.height=(parseInt(obj.style.height)+8).toString()+"px";
}
if(str=="down")
{
if (parseInt(obj.style.height)<8)
{ clearInterval(handle);
   obj.style.display="none";
}
else 
   obj.style.height=(parseInt(obj.style.height)-8).toString()+"px"; 
}
}

</script>
</HEAD>
<BODY onLoad="start();"><!-- 内容区 -->
<%
sqlCG = "select distinct s1.billcode,custname,depotname,memo,username,adddate from ((select billcode+goodscode as bg,goodscode,sum(number) as t_num,billcode,custname,depotname,memo,username,adddate from s_billdetail where billtype = '采购订货' and s_billdetail.check=1 group by billcode,goodscode,custname,depotname,memo,username,adddate) as s1 left join (select planbillcode+goodscode as pg,planbillcode,goodscode,sum(number) as t_num from s_billdetail group by planbillcode,goodscode) as s2 on s1.bg = s2.pg) where s1.t_num > s2.t_num or s2.t_num is null"
set rsCG = Server.CreateObject("adodb.recordset")
rsCG.open sqlCG, conn, 1, 1

sqlXS = "select distinct s1.billcode,custname,depotname,memo,username,adddate from ((select billcode+goodscode as bg,goodscode,sum(number) as t_num,billcode,custname,depotname,memo,username,adddate from s_billdetail where billtype = '销售订货' and s_billdetail.check=1 group by billcode,goodscode,custname,depotname,memo,username,adddate) as s1 left join (select planbillcode+goodscode as pg,planbillcode,goodscode,sum(number) as t_num from s_billdetail group by planbillcode,goodscode) as s2 on s1.bg = s2.pg) where s1.t_num > s2.t_num or s2.t_num is null"
set rsXS = Server.CreateObject("adodb.recordset")
rsXS.open sqlXS, conn, 1, 1

sqlFK = "select s1.custname,(s1.tmon+s3.tmon) as totalmoney,s2.tmon as paymoney,s1.tmon+s3.tmon-s2.tmon as shouldmoney from (((select case when sum(money*flag) is null then 0 else sum(money*flag) end as tmon,custname from s_billdetail where billtype like '采购%' and flag<>0 group by custname) as s1 left join (select case when sum(money*sign*-1) is null then 0 else sum(money*sign*-1) end as tmon,custname from t_cash where cashtype='付款' group by custname) as s2 on s2.custname = s1.custname) left join (select startmoneyf as tmon,custname from t_custom) as s3 on s3.custname = s1.custname) where s1.tmon+s3.tmon-s2.tmon > 0"
set rsFK = Server.CreateObject("adodb.recordset")
rsFK.open sqlFK, conn, 1, 1

sqlSK = "select s1.custname,(s1.tmon+s3.tmon) as totalmoney,s2.tmon as paymoney,s1.tmon+s3.tmon-s2.tmon as shouldmoney from (((select case when sum(money*flag) is null then 0 else sum(money*flag) end as tmon,custname from s_billdetail where billtype like '销售%' and flag<>0 group by custname) as s1 left join (select case when sum(money*sign*-1) is null then 0 else sum(money*sign*-1) end as tmon,custname from t_cash where cashtype='付款' group by custname) as s2 on s2.custname = s1.custname) left join (select startmoneyf as tmon,custname from t_custom) as s3 on s3.custname = s1.custname) where s1.tmon+s3.tmon-s2.tmon > 0"
set rsSK = Server.CreateObject("adodb.recordset")
rsSK.open sqlSK, conn, 1, 1
%>
<div id="tip">
<h1><a href="javascript:void(0)" onClick="end()">×</a>系统提示</h1>
<P>还有<font color="#FF0000"><%=rsCG.recordcount%></font>张采购订单未完结<a href="../bills/addbill.asp?type=CG">马上处理>></a></P>
<p>还有<font color="#FF0000"><%=rsXS.recordcount%></font>张销售订单未完结<a href="../bills/addbill.asp?type=XS">马上处理>></a></p>
<p>还有<font color="#FF0000"><%=rsFK.recordcount%></font>供应商应付款未处理<a href="../cash/addcash.asp?add=true&type=fk">马上处理>></a></p>
<p>还有<font color="#FF0000"><%=rsSK.recordcount%></font>客户应收款未处理<a href="../cash/addcash.asp?add=true&type=sk">马上处理>></a></p>
</div>
<br>
<br>
<table align="center" width="673">
<tr>
	<td width="182">
    	<table border="0" cellpadding="0" cellspacing="0">
        	<tr>
        	  <td colspan="2">
              <img src="../img/main_3.gif" align="absmiddle">
              </td>
       	  </tr>
        	<tr>
        	  <td colspan="2"><img src="../img/main1_jxc_14.gif"></td>
           	</tr>
        	<tr>
        	  <td valign="top">
              <a href="../cash/addcash.asp?type=fk&add=true"><img src="../img/main1_19.gif" alt="" name="fk" align="absmiddle" id="fk" border="0"></a>
              </td>
        	  <td rowspan="2" valign="top"><img src="../img/main1_jxc_20.gif"></td>
      	  </tr>
        	<tr>
        	  <td valign="top"><img src="../img/main1_jxc_30.gif"></td>
       	  </tr>
        	<tr>
        	  <td colspan="2"><img src="../img/main_33.gif" align="absmiddle"></td>
           	</tr>
      </table>           	
    </td>
    <td width="96">
<table border="0" cellpadding="0" cellspacing="5">
        	<tr>
        	  <td width="143">
              <a href="../Bills/addbill.asp?type=CG"><img src="../img/main1_4.gif" name="cg" id="cg" border="0"></a>
              </td>
      	  </tr>
        	<tr>
            	<td><img src="../img/main1_jxc_10.gif"></td>
            </tr>
        	<tr>
        	  <td><a href="../Bills/inbill.asp"><img src="../img/main1_12.gif" name="rk" id="rk" border="0"></a></td>
      	  </tr>
        	<tr>
            	<td><img src="../img/main1_jxc_24.gif"></td>
            </tr>
        	<tr>
            	<td><a href="../Bills/pandianbill.asp"><img src="../img/main1_28.gif" name="pd" id="pd" border="0"></a></td>
            </tr>
        	<tr>
            	<td><img src="../img/main1_jxc_37.gif"></td>
            </tr> 
        	<tr>
            	<td><a href="../Bills/selectbackbill.asp?type=CT"><img src="../img/main1_39.gif" name="ct" id="ct" border="0"></a></td>
            </tr>
 </table>  
</td>
    <td width="252">
    	<table>
        	<tr>
            	<td width="252">
                	<table width="252" border="0" cellpadding="0" cellspacing="0">
                    	<tr>
                        	<td><img src="../img/main_5.gif" width="125" height="113"></td>
                           	<td><img src="../img/main_6.gif" width="126" height="113"></td>
                        </tr>
                    </table>
                	<table width="252" border="0" cellpadding="0" cellspacing="0">
                    	<tr>
                       	  <td width="83">
                            	<table border="0" cellpadding="0" cellspacing="0">
                                	<tr>
                                    	<td>
                                        	<img src="../img/main_15.gif"></td>
                                    </tr>
                                    <tr>
                                    	<td>
                                       		 <img src="../img/main_26.gif" width="83" height="63"></td>
                                  </tr>
                                </table>
                       	  </td>
                            <td>
               	  <table border="0" cellpadding="0" cellspacing="0">
                                	<tr>
                                    	<td>
                                        <img src="../img/main1_jxc_16.gif" width="84" height="22"></td>
                    </tr>
                                	<tr>
                                    	<td><a href="../report/depotcount.asp"><img src="../img/main1_21.gif" name="depot" width="84" height="83" align="absmiddle" id="depot" border="0"></a></td>
                    </tr>
                                	<tr>
                                    	<td>
                                        <img src="../img/main1_jxc_31.gif" width="84" height="21"></td>
                    </tr>
                              </table>
                          </td>
                       	  <td width="84">
                           		<table border="0" cellpadding="0" cellspacing="0">
                                	<tr>
                                    	<td>
                                        	<img src="../img/main_17.gif" width="84" height="63"></td>
                                    </tr>
                                    <tr>
                                    	<td>
                                        	<img src="../img/main_27.gif" width="84" height="63"></td>
                                    </tr>
                                </table>    
                          </td>
                      </tr>
                  </table>
                	<table border="0" cellpadding="0" cellspacing="0">
                    	<tr>
                        	<td><img src="../img/main_34.gif" width="125" height="113"></td>
                     	<td>
                          <img src="../img/main_35.gif" width="126" height="113"></td>
                        </tr>
                    </table></td>
            </tr>
       	</table>
    </td>
    <td width="96">
<table border="0" cellpadding="0" cellspacing="5">
        	<tr>
        	  <td width="143"><a href="../Bills/addbill.asp?type=XS"><img src="../img/main1_7.gif" name="xs" id="xs" border="0"></a></td>
      	  </tr>
        	<tr>
            	<td><img src="../img/main1_jxc_10.gif"></td>
            </tr>
        	<tr>
        	  <td><a href="../Bills/outbill.asp"><img src="../img/main1_29.gif" name="ck" id="ck" border="0"></a></td>
      	  </tr>
        	<tr>
            	<td><img src="../img/main1_jxc_24.gif"></td>
            </tr>
        	<tr>
           	  <td><a href="../Bills/diaobobill.asp"><img src="../img/main1_13.gif" name="db" id="db" border="0"></a></td>
            </tr>
        	<tr>
            	<td><img src="../img/main1_jxc_37.gif"></td>
            </tr> 
        	<tr>
           	  <td><a href="../Bills/selectbackbill.asp?type=XT"><img src="../img/main1_40.gif" name="xt" id="xt" border="0"></a></td>
            </tr>
 </table>  
</td>
	<td width="182">
    	<table border="0" cellpadding="0" cellspacing="0">
        	<tr>
        	  <td colspan="2"><img src="../img/main_8.gif" align="absmiddle"></td>
       	  </tr>
        	<tr>
        	  <td colspan="2"><img src="../img/main1_jxc_18.gif"></td>
           	</tr>
        	<tr>
        	  <td rowspan="2" valign="top"><img src="../img/main1_jxc_20.gif"></td>
        	  <td valign="top"><a href="../cash/addcash.asp?type=sk&add=true"><img src="../img/main1_23.gif" alt="" name="sk" align="absmiddle" id="sk" border="0"></a></td>
      	  </tr>
        	<tr>
        	  <td valign="top"><img src="../img/main1_jxc_30.gif"></td>
      	  </tr>
        	<tr>
        	  <td colspan="2"><img src="../img/main_36.gif" align="absmiddle"></td>
           	</tr>
      </table>           	
    </td>
</tr>
</table>
<table align="center" width="673">
<tr>
	<td align="right"><a href="goods.asp"><img id="goods" border="0" src="../img/icon_gys_da.gif"></a></td>
    <td align="center"><a href="employee.asp"><img id="employee" border="0" src="../img/icon_gys_da3.gif"></a></td>
    <td align="left"><a href="custom.asp"><img id="cust" border="0" src="../img/icon_gys_da2.gif"></a></td>
</tr>
</table>

</BODY>
</HTML>

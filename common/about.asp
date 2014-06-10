<!-- #include file="../inc/conn.asp" -->
<html>
<head>
<title>关于</title>
<link rel="stylesheet" href="../style.css" type="text/css">
<style type="text/css">
<!--
body {
  border:0;
  overflow:	hidden;
  margin: 0;
  font:	MessageBox;
  background-color:#D6D3CE;
}
-->
</style>
</head>
<%
sql = "select * from t_softinfo"
Set rs = conn.Execute(sql)
%>
<body oncontextmenu="return false" topmargin="0">
<table width="450" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td><img src="../img/aboutus.jpg"></td>
</tr>
<tr>
    <td> <fieldset style="padding-bottom:0px"><legend>关于我们</legend>
      <table border="0" cellspacing="5" cellpadding="5" width="95%" align="center">
        <tr> 
          <td align="left"><font color="#0000FF">小二科技进销存管理系统B/S版&nbsp;V<%=rs("Version")%>&nbsp;Build20120417<br>数据库版本号：<%=rs("DBVersion")%></font></td>
        </tr>
		
        <tr> 
          <td width="100%">官方网站：<a href="http://www.hokilly.com" target="_blank">http://www.hokilly.com</a> 
            <br>
			发布时间：2012-10-26
			<br>
            电子信箱：<a href="mailto:Hokilly@126.com">Hokilly@126.com</a> 
            <br>
            电话传真：0519-85355013
			<br>
			技 术Q Q：15916190
          </td>
        </tr>
        
      </table>
 </fieldset>
 </td>
</tr>
<tr>
    <td align="right" valign="top">
  <!-------  <input type="button" class="button" value=" 确定 "   onMouseOver="this.className='button_over';" onMouseOut="this.className='button';" onClick="window.close()"> ----->
    </td>
</tr>
</table>
<%
close_rs(rs)
endconnection
%>
</body>
</html>

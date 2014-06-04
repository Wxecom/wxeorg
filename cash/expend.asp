<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<% call CheckAuthority("strBrowseZC") %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<title>其他支出</title>
<script>
function check_edit(str){
window.open ('add_cashbank.asp?add=false&type=ZC&billcode='+str, 'othercashbill', 'height=300, width=650,top=100,left=150,toolbar=no,menubar=no, scrollbars=yes, resizable=no,location=no, status=no');
}

function NewBill(){
window.open ('add_cashbank.asp?add=true&type=ZC', 'othercashbill', 'height=300, width=650,top=100,left=150,toolbar=no,menubar=no, scrollbars=yes, resizable=no,location=no, status=no');
}

function delall(){
	var iCount = 0;
	var index = 0;
	var iSucceed = 0;
	var errMsg = "";
	$("._del").each(
		function(i)
		{
			chk = $(this)
			if(chk.attr("checked") == true){
				iCount = iCount + 1;
			}
		}
	);
	if(confirm('确定要删除选择的'+iCount+'条单据吗?')){
		$("._del").each( 
			function(i)
			{
				chk = $(this)
				if(chk.attr("checked") == true){
					$.post("../select/delallcashbank.asp",{billcode:escape(chk.val())},
					function(data)
					{ 
						index = index + 1;
						if(data == "True"){
							iSucceed = iSucceed + 1;
						}else{
							errMsg = errMsg + data + "\n";
						}
						if(index == iCount){
							alert("成功删除"+iSucceed+"张单据!");
							if(iCount != iSucceed){
								alert("未成功删除单据：\n" + errMsg);
							}
							window.location.reload();
						}
					}
					);
				}
			}
		)//遍历结束
	}
}
</script>
</head>

<body style="background:#FFFFFF">
<div style="background:url(../images/main-02.gif); width:100%; height:20; padding-left:0" ><img src="../images/main-01.gif"></div>
<div style="padding-left:10px">
<%
if Request.Form("date1") = "" then
  s_date1 = Year(Date()) & "-" & Month(Date()) & "-1"
else
  s_date1 = Request.Form("date1")
end if

if Request.Form("date2") = "" then
  s_date2 = formatdate(date)
else
  s_date2 = Request.Form("date2")
end if
%>
<br>
<div align="left" class="STYLE1">其他支出</div>
<table align="left" border="0"><tr><td>
<table border="0">
<tr>
  <td>
    <form name="form1" action="expend.asp" method="post">
	<table align="left">
	<tr>
	  <td>制单日期:<input type="text" name="date1" id="date1" size="16" value="<%=s_date1%>"><%showdate("date1")%>
	               <input type="text" name="date2" id="date2" size="16" value="<%=s_date2%>"><%showdate("date2")%>
	  </td>
	  <td>项目名称:<input type="text" name="projectname" value="<%= Request.Form("projectname") %>">
	  </td>
	  <td>经办人:<select name="user"><option value=""></option>
	                         <% sql = "select * from t_Employee"
                             Set a = conn.Execute(sql)
                             Do While a.EOF = False
                              s_name = a("name")
							  if Request.Form("user") = s_name then
                              Response.Write "<option value="&s_name&" selected>"&s_name&"</option>"
							  else
							  Response.Write "<option value="&s_name&">"&s_name&"</option>"
							  end if
                              a.movenext
                             Loop
							 close_rs(a)%></select></td>
	  <td valign="top"><input type="submit" class="button" onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  value=" 查 询 "></td>	  
	</tr>
	</table>
	</form>
  </td>
</tr>
</table>
<hr>	
<table align="left" border="0">
<tr><td>
	<input type="button" class="button but_mar" value="新增" onClick="NewBill();">
	<input type="button" class="button but_mar"  value="删除" onClick="delall();">
</td></tr>
<tr valign="top">
<td>
	<%
	if Request.Form("projectname") = "" then
		s_projectname = ""
	else
		s_projectname = " and projectname like '%" & Request.Form("projectname") & "%'"
	end if
	
	if Request.Form("user") = "" then
		s_user = ""
	else
		s_user = " and user = '" & Request.Form("user") & "'"
	end if
	sql = "select CONCAT('<a href=# onClick=check_edit(''',billcode,''')>',billcode,'</a>') as abillcode,t_cashbank.* from t_cashbank where billtype = '其他支出' and (AddDate)<='"&s_date2&"' and (AddDate)>='"&s_date1&"'" & s_projectname & s_user
	call showpage(sql,"income",1)
	endconnection
	%>
</td>
</tr>
</table>
</div>
</body>
</html>

<%@LANGUAGE="VBSCRIPT"%>
<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<html>
<head>
<title>仓库资料</title>
</head>
<body bgcolor="#F3F1E9">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<script>
function save(){
	if (document.depot.depotname.value=="")
	{
		alert('请填写仓库名称！');
		document.depot.depotname.focus();
		return; 
	}
	if("<%=request.querystring("type")%>"=="add"){
		checkvalue();
	}
	document.depot.submit();
}

function checkvalue(){
	var obj = document.all.user_qx;
	var userlist="";
	for(var i=0;i<obj.length;i++){
		if (obj[i].checked){
			userlist=userlist + obj[i].value + ",";
		}
	}
	document.depot.userlist.value=userlist;
}

function check_name(){
	if ($("#depotname").val() == "") {
		$("#div_check_name").attr("innerHTML","<li style='float:left;'><img src='../img/btn_false.gif'></li><li style='float:left;width:auto'><strong><font color=#FF0000>不能为空</font></strong></li>");
	}else{
    $.post("../select/adddepot_ajax_name.asp",{depotname:escape($("#depotname").val())},
	function(data)
	{ 
       if (data == "True"){
		   $("#div_check_name").attr("innerHTML","<img src='../img/btn_true.gif'>"); 
		   }else{
		   $("#div_check_name").attr("innerHTML","<li style='float:left;'><img src='../img/btn_false.gif'></li><li style='float:left;width:auto'><strong><font color=#FF0000>名称重复</font></strong></li>"); 
	   }
	}
	);	
	}
}
</script>
<br>
<div align="center"><span class="STYLE1">仓库资料</span></div>
<br>
<%
if Request.QueryString("depotname")<>"" then
	sqlDepot = "select * from t_depot where depotname = '" & Request.QueryString("depotname") & "'"
	set rsDepot = Server.CreateObject("adodb.recordset")
	rsDepot.open sqlDepot,conn,1,1
	s_id = rsDepot("id")
	s_depotname = rsDepot("depotname")
	s_department = rsDepot("department")
	s_controller =  rsDepot("controller")
	s_principal = rsDepot("principal")
	s_tel = rsDepot("tel")
	s_address = rsDepot("address")
	s_remark = rsDepot("remark")
end if
%>
<form name="depot" method="post" action="../action/savedepot.asp?type=<%=request.querystring("type")%>&depotid=<%=s_id%>">
<table width="600" style="margin-left:280px" border="0" align="center">
  <tr>
    <td width="100" align="right">仓库名称:</td>
    <td width="200" align="left">
    <input type="text" id="depotname" maxlength="50" name="depotname" <%if request.querystring("type")="edit" then response.write "readonly" else response.write "onBlur=check_name()" end if %> size="20" value="<%=s_depotname%>"><font color="#FF0000">*</font><div id="div_check_name"></div></td>
    <td width="300"></td>
  </tr>
  <tr>
  	<td align="right">所属部门:</td>
  	<td width="200" align="left" colspan="2"><select name="department" id="department">
  			<%sql_dep = "select name from t_department order by id"
  				set rs_dep = server.createobject("adodb.recordset")
  				rs_dep.open sql_dep,conn,1,1
  				for i=0 to rs_dep.recordcount - 1
  					if rs_dep.bof or rs_dep.eof then exit for
  					response.write "<option value="""&rs_dep("name")&""" "
  					if rs_dep("name") = s_department then
  						response.write "selected"
  					end if
  					response.write " >"&rs_dep("name")&"</option>"
  					rs_dep.movenext
  				next
  				close_rs(rs_dep)
  			%>
  		</select>
  	</td>
  </tr>
  <tr>
    <td align="right">资产管理员:</td>
    <td align="left"><select name="controller">
    	<%sql_emp="select name from t_employee order by id"
    		set rs_emp = server.createobject("adodb.recordset")
    		rs_emp.open sql_emp,conn,1,1
    		rs_emp.movefirst
    		for i=0 to rs_emp.recordcount - 1
					if rs_emp.bof or rs_emp.eof then exit for
					response.write "<option value="""&rs_emp("name")&""" "
					if rs_emp("name") = s_controller then
					  response.write "selected"
					end if
					response.write " >"&rs_emp("name")&"</option>"
					rs_emp.movenext
				next
    	%>
    </select>
    </td>
    <td></td>
  </tr>
  <tr>
    <td align="right">负&nbsp;责&nbsp;人:</td>
    <td align="left"><select name="principal">
    								<%rs_emp.movefirst
    									for i=0 to rs_emp.recordcount - 1
						  					if rs_emp.bof or rs_emp.eof then exit for
						  					response.write "<option value="""&rs_emp("name")&""" "
						  					if rs_emp("name") = s_principal then
						  						response.write "selected"
						  					end if
						  					response.write " >"&rs_emp("name")&"</option>"
						  					rs_emp.movenext
					  					next
					  					close_rs(rs_emp)
    								%>
    								</select>
    </td>
    <td></td>
   </tr>
   <tr> 
    <td align="right">电&nbsp;&nbsp;&nbsp;&nbsp;话:</td>
    <td align="left"><input name="tel" type="text" maxlength="50" size="20" value="<%=s_tel%>"></td>
    <td></td>
  </tr>
  <tr>
    <td align="right">地&nbsp;&nbsp;&nbsp;&nbsp;址:</td>
    <td colspan="2" align="left"><input name="address" type="text" size="64" value="<%=s_address%>"></td>
    <td></td>
  </tr>
  <tr>
    <td align="right">备&nbsp;&nbsp;&nbsp;&nbsp;注:</td>
    <td colspan="2" align="left"><input type="text" name="remark" maxlength="200" size="64" value="<%=s_remark%>"></td>
    <td></td>
  </tr>
  <%if request.querystring("type")="add" then%>
  <tr>
    <td align="right">设置管理权限:</td>
    <td colspan="2" align="left"><label>
      <input <%If (CStr(userlist) = CStr(userlist)) Then Response.Write("checked=""checked""") : Response.Write("")%> type="checkbox" name="userlist" id="userlist">
    </label></td>
    <td></td>
  </tr>
  <%end if%>
  <tr>
    <td colspan="3" align="center">
  
    	<input type="button" class="button" value=" 保 存 " onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  onClick="save()">	
        
       	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <input type="reset" class="button" value=" 重 置 " onMouseOver="this.className='button_over';" onMouseOut="this.className='button';" ></td>
  </tr>
</table>
</form>
<%
endconnection
%>
</body>
</html>

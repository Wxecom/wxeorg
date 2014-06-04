<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="../inc/config.asp" -->
<% call CheckAuthority("strBrowseCustom") %>
<html>
<head>
<title>往来单位资料</title>
</head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="../style.css" type="text/css">
<link rel="stylesheet" href="dtree.css" type="text/css">
<script type="text/javascript" src="../js/dtree.js"></script>
<script language=javascript>
function clicknode(nodename,nodecode){
	if (document.getElementById("nodename")!=null){
		document.getElementById("nodename").value = String(nodename);	
	}
	if (document.getElementById("typecode")!=null){
		document.getElementById("typecode").value = String(nodecode);	
	}	
	closeProc();
}

function edit(custname){
	window.open ('addcust.asp?add=false&custname='+custname, 'addcust', 'height=600, width=650,top=100,left=150,toolbar=no,menubar=no, scrollbars=yes, resizable=no,location=no, status=yes')
}

function del(){
  if (document.getElementById('temp').value==""){
    alert("请选择往来单位！");
    return false; 
	}
	if(!confirm('确定要删除该往来单位的资料吗?')){
		return false
	}else{
		window.location.href='../action/delcust.asp?id='+document.getElementById('temp').value
	}
}

function showProc(){
  message_box.style.visibility='visible';
  //创建灰色背景层
  procbg = document.createElement("div"); 
  procbg.setAttribute("id","mybg"); 
  procbg.style.background = "#f7f7f7"; 
  procbg.style.width = "100%"; 
  procbg.style.height = "100%"; 
  procbg.style.position = "absolute"; 
  procbg.style.top = "0"; 
  procbg.style.left = "0"; 
  procbg.style.zIndex = "500"; 
  procbg.style.opacity = "0.3"; 
  procbg.style.filter = "Alpha(opacity=30)"; 
  //背景层加入页面
  document.body.appendChild(procbg);
  document.body.style.overflow = "hidden";
}
//拖动
function drag(obj){  
     var s = obj.style;  
     var b = document.body;   
  var x = event.clientX + b.scrollLeft - s.pixelLeft;   
  var y = event.clientY + b.scrollTop - s.pixelTop; 
 
  var m = function(){  
   if(event.button == 1){  
    s.pixelLeft = event.clientX + b.scrollLeft - x;   
    s.pixelTop = event.clientY + b.scrollTop - y;   
   }else {
    document.detachEvent("onmousemove", m);
   }  
  }  
  document.attachEvent("onmousemove", m);
  if(!this.z) 
   this.z = 999;   
  s.zIndex = ++this.z;   
  event.cancelBubble = true;   
}
 
function closeProc(){
  message_box.style.visibility= "hidden";
  procbg.style.visibility = "hidden";
}
function newcust(){
	openwin('addcust.asp?add=true',650,600);
}
function delall(){
	var iCount = 0;
	var index = "";
	$("._del").each(
		function(i)
		{
			chk = $(this)
			if(chk.attr("checked") == true){
				iCount = iCount + 1;
			}
		}
	);
	if(confirm('确定要删除选择的'+iCount+'条往来单位的资料吗?')){
		$("._del").each( 
			function(i)
			{
				chk = $(this)
				if(chk.attr("checked") == true){
					index = chk.attr("id");
					index = index.substring(2);
					$("#tr" + index).remove();
					$.post("../select/delallcust.asp",{custname:escape(chk.val())},
					function(data)
					{ 
					}
					);
				}
			}
		)//遍历结束
	}
}
</script>
<body style="background:#FFFFFF">
<div style="background:url(../images/main-02.gif); width:100%; height:20; padding-left:0" ><img src="../images/main-01.gif"></div>
<div style="padding-left:10px">
<br>
<div align="left" class="style1">往来单位资料</div>
<form name="form1" method="post" action="custom.asp">
<table border="0">
<tr>
	<td colspan="2">
往来单位名称：<input type="text" id="custname" name="custname" size="16" value="<%=request.Form("custname")%>">
往来单位类别：<input type="text" id="nodename" name="nodename" size="16" value="<%=request.Form("nodename")%>">
<a href="#SelectDate" onClick="showProc()"><img border="0" src="../img/choose.gif" width="21" height="17"></a><input type="hidden" name="typecode" value="<%=request.Form("typecode")%>"></td>
<td valign="top">
	<input type="submit" class="button" onMouseOver="this.className='button_over';" onMouseOut="this.className='button';"  value="查询">	
	<input type="button" value="新增"  onMouseOver="this.className='button_over';" onMouseOut="this.className='button';" onClick="newcust();" class="button">
	<input type="button" value="删除"  onMouseOver="this.className='button_over';" onMouseOut="this.className='button';" onClick="delall()" class="button">
</td>
</tr>
</form>
</table>

<hr>
<table>
<tr>



<%
sql_cust = "select custname from t_user where username='"&request.cookies("username")&"'"
Set rs_cust = server.createobject("adodb.recordset")
rs_cust.open sql_cust,conn,1,1
if rs_cust("custname")<>"" then
	arr = split(rs_cust("custname"),",")
	temp_cust = ""
	if ubound(arr) <> -1 then
		for i = lbound(arr) to ubound(arr)-1
		  temp_cust = temp_cust & "'" & arr(i) & "',"
		next
	end if
end if

if (request.cookies("username") <> "admin") then
  s_cust = " "
else
  s_cust = " and custname like  '%"&request("custname")&"%' "
end if
If request("custname") = "" Then
    If request("nodename") = "" Then
        sql = "select CONCAT('<a href=# onClick=edit(''',custname,''')>',custname,'</a>') as aCustname,t_custom.* from t_custom where 1=1" & s_cust&" order by CustID desc"
    Else
        sql = "select CONCAT('<a href=# onClick=edit(''',custname,''')>',custname,'</a>') as aCustname,t_custom.* from t_custom where 1=1 and code like '"&request("typecode")&"%'" & s_cust&" order by CustID desc"
    End If
Else
    If request("nodename") = "" Then
        sql = "select CONCAT('<a href=# onClick=edit(''',custname,''')>',custname,'</a>') as aCustname,t_custom.* from t_custom where 1=1 and CustName like '%"&request("custname")&"%'" & s_cust&" order by CustID desc"
    Else
        sql = "select CONCAT('<a href=# onClick=edit(''',custname,''')>',custname,'</a>') as aCustname,t_custom.* from t_custom where 1=1 and code like '"&request("typecode")&"%' and CustName='"&request("custname")&"'" & s_cust&" order by CustID desc"
    End If
End If
call showpage(sql,"custom",1)
%>

<div id="container"></div>
<div align="center"><span class="style1"></span></div>
<div id="message_box" style="position:absolute;left:50%;top:5%;width:250px;height:400px;filter:dropshadow(color=#666666,offx=3,offy=3,positive=2);z-index:1000;visibility:hidden">
   <div id= "message" style="border:#036 solid; border-width:1 1 3 1;width:95%; height:95%; background:#fff; color:#036; font-size:12px; line-height:150%;">
    <!-- DIV弹出状态行：标题、关闭按钮 -->
    <div style="background:#7DC6F0;height:5%;font-family:宋体; font-size:13px; padding:3 5 0 5; color:#00385E" onMouseDown="drag(message_box);return false">
     <span style="display:inline;width:100%;position:absolute;font-size:120%">单击选择类别</span>
     <span onClick="closeProc();" style="float:right;display:inline;cursor:pointer;font-size:150%">×</span>
    </div>
<script language="JavaScript">
	d = new dTree('d');
	<%
    Response.Write "d.add(0,-1,'往来单位分类','edittree.asp?type=cust&name="&s_name&"&id="&i_id&"&pid=0');"
		sql_gtype = "select * from t_tree where type='cust'"
		set rs_gtype = conn.Execute(sql_gtype)
		do while rs_gtype.eof=false
		  i_id=rs_gtype("id")
			i_pid=rs_gtype("pid")
			s_name=rs_gtype("name")
			s_url=rs_gtype("url")
			s_code=rs_gtype("code")
		  Response.Write "d.add("&i_id&","&i_pid&",'"&s_name&"',""javascript:clicknode('"&s_name&"','"&s_code&"')"",'','','','','','"&s_code&"');"
		  rs_gtype.movenext
		loop
		close_rs(rs_gtype)
		%>
	document.write(d);
</script>
</table>
</div><!-- message -->
</div><!-- message_box -->
<%endconnection%>
</div>
</body>
</html>
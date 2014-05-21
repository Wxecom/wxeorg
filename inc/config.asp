<!-- #include file="conn.asp" -->
<!-- #include file="../action/checkuser.asp" -->
<!-- #include file="page.asp" -->
<%
Response.ContentType = "text/html"
Response.Charset = "GB2312"   '解决乱码问题
%>
<script language="JavaScript" src="../js/calendar.js"></script>
<script language="JavaScript" src="../js/xmlhttp.js"></script>
<script language="JavaScript" src="../js/showphoto.js"></script>
<script language="JavaScript" src="../js/addrow.js"></script>
<script language="JavaScript" src="../js/jquery.js"></script>
<SCRIPT>

function openwin(URL,x,y){
var URL;
var x1=window.screen.width;
var y1=window.screen.height;
x2=(x1-x)/2;
y2=(y1-y)/2;
window.open(URL,'','top='+y2+',left='+x2+',width='+x+',height='+y+',status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes')
}

function setHiddenCol1(oTable,iCol)
{
    for (i=0;i < oTable.rows.length ; i++)
    {
        oTable.rows[i].cells[iCol].style.display = oTable.rows[i].cells[iCol].style.display=="none"?"block":"none";
		}
}


function change(){
	var   oObj   =   event.srcElement;
	if(oObj.tagName.toLowerCase() == "td"){
		var   oTr   =   oObj.parentNode;
		for(var   i=1;   i<$("#tbl").find("tr").length;   i++){
			$("#tbl").find("tr").css("backgroundColor","#ffffff");
		}
		oTr.style.backgroundColor   =   "#CCCCFF";   
		oTr.tag   =   true;
	}
}


function   out(){
	var   oObj   =   event.srcElement;
	if(oObj.tagName.toLowerCase()   ==   "td") {
		var   oTr   =   oObj.parentNode;
		if(!oTr.tag)   oTr.style.backgroundColor   =   "";
	}
}

function   over()
{
var   oObj   =   event.srcElement;
if(oObj.tagName.toLowerCase()   ==   "td")
{
var   oTr   =   oObj.parentNode;
if(!oTr.tag)   oTr.style.backgroundColor   =   "#E1E9FD";
}


}

function addtext(i_row,str)
{
document.getElementById("temp").value=str;
}

function GetDIV(oobj,val_i,str)
{

	
if (oobj == 'block')
{
var   oObj   =   event.srcElement;
if(oObj.tagName.toLowerCase()   ==   "td")
{
var   oTr   =   oObj.parentNode;
if(!oTr.tag)   oTr.style.backgroundColor   =   "#E1E9FD";
}
}
else
{
var   oObj   =   event.srcElement;
if(oObj.tagName.toLowerCase()   ==   "td") 
{
var   oTr   =   oObj.parentNode;
if(!oTr.tag)   oTr.style.backgroundColor   =   "";
}
}
 document.getElementById("search_suggest").style.display=oobj;
 document.getElementById("temp1").value = str;
 document.getElementById("search_suggest").style.top=event.clientY+20+document.body.scrollTop;
 document.getElementById("search_suggest").style.left=event.clientX+20+document.body.scrollLeft;
}

function CreateReturnValue(closeWindow,i){
	var tabindex=self.opener.document.getElementById('sqd').rows.length
	j=Number(self.opener.document.getElementById("rowcount").value) + 1
	self.opener.document.getElementById("rowcount").value=j
	var T1=""+(tabindex)+"";
	var T2="<a href='#' onClick='del_row("+j+")'>删除</a>";
	var T3="<input name=\"goodscode\" type=\"text\" class=INPUT1 id=goodscode"+j+" size=16 />"
	var T4="<input name=goodsname type=text class=INPUT1 id=goodsname"+j+"  size=16 />"
	var T5="<input name=goodsunit type=text class=INPUT1 id=goodsunit"+j+"  size=13 />" 
	var T6="<input name=units type=text class=INPUT1 id=units"+j+"  size=5 />"
	var T7="<input onkeyup=count("+j+") name=price  type=text class=INPUT1 id=price"+j+"  size=8 />"
	var T8="<input onkeyup=count("+j+") name=number type=text class='number' id=number"+j+"  size=8 />"
	var T9="<input name=money type=text id=money"+j+" class='money' size=8 readonly />"
	var T10="<input name=remark type=text class=INPUT1 id=remark"+j+" size=10 /><input name=aveprice type=hidden class=INPUT1 id=aveprice"+j+"  size=20 /><input name=fact_num type=hidden id=fact_num"+j+" />"
	var objRow = self.opener.document.all.sqd.insertRow(-1);
	var objCel = objRow.insertCell(0);
	objCel.innerHTML =T1;
	var objCel = objRow.insertCell(1);
	objCel.innerHTML =T2 ;
	var objCel = objRow.insertCell(2);
	objCel.innerHTML =T3 ;
	var objCel = objRow.insertCell(3);
	objCel.innerHTML =T4 ;
	var objCel = objRow.insertCell(4);
	objCel.innerHTML =T5 ;
	var objCel = objRow.insertCell(5);
	objCel.innerHTML =T6 ;
	var objCel = objRow.insertCell(6);
	objCel.innerHTML =T7 ;
	var objCel = objRow.insertCell(7);
	objCel.innerHTML =T8 ;
	var objCel = objRow.insertCell(8);
	objCel.innerHTML =T9 ;
	var objCel = objRow.insertCell(9);
	objCel.innerHTML =T10 ;

	var goodscode=self.opener.document.getElementById("goodscode"+j);
	var goodsname=self.opener.document.getElementById("goodsname"+j);
	var goodsunit=self.opener.document.getElementById("goodsunit"+j);
	var units=self.opener.document.getElementById("units"+j)
	var price=self.opener.document.getElementById("price"+j)
	var number=self.opener.document.getElementById("number"+j)
	var money=self.opener.document.getElementById("money"+j)
	var remark=self.opener.document.getElementById("remark"+j);
	var aveprice=self.opener.document.getElementById("aveprice"+j);
	var fact_number=self.opener.document.getElementById("fact_num"+j);

	goodscode.value=document.getElementById('tbl').rows[i].cells[1].innerHTML;
	goodsname.value=document.getElementById('tbl').rows[i].cells[2].innerHTML;
	goodsunit.value=document.getElementById('tbl').rows[i].cells[3].innerHTML;
	units.value=document.getElementById('tbl').rows[i].cells[4].innerHTML
	number.value=1;
	remark.value=document.getElementById('tbl').rows[i].cells[7].innerHTML;
	aveprice.value=document.getElementById('tbl').rows[i].cells[9].innerHTML;
	fact_number.value=document.getElementById('tbl').rows[i].cells[8].innerHTML;
	if ( self.opener.document.getElementById("total_number").value=="")
	{
		self.opener.document.getElementById("total_number").value=0
	}
	if ( self.opener.document.getElementById("total_money").value=="")
	{
		self.opener.document.getElementById("total_money").value=0
	}
	self.opener.document.getElementById("total_number").value=parseFloat(self.opener.document.getElementById("total_number").value)+ 1;
	
	<%
	set rsMemory = Server.CreateObject("adodb.recordset")
	sql = "select memoryprice from t_softinfo"
	rsMemory.open conn,1,1
	f_memoryprice = rsMemory(0)
	%>
	
	if ("<%response.write request.QueryString("type")%>"=="XS")
	{
		if ("<%response.Write f_memoryprice%>" == "False")
		{
			price.value=document.getElementById('tbl').rows[i].cells[6].innerHTML;
		}else{
			price.value=document.getElementById('tbl').rows[i].cells[10].innerHTML;	
		}
		self.opener.document.getElementById("total_money").value=parseFloat(self.opener.document.getElementById("total_money").value)+parseFloat(price.value)*parseFloat(number.value);	
		self.opener.document.getElementById("zdprice").value=self.opener.document.getElementById("total_money").value;
		self.opener.document.getElementById("yfprice").value=self.opener.document.getElementById("total_money").value;
	}
	else if ("<%response.write request.QueryString("type")%>"=="CG")
	{	
		if ("<%response.Write f_memoryprice%>" == "False")
		{
			price.value=document.getElementById('tbl').rows[i].cells[5].innerHTML;
		}else{
			price.value=document.getElementById('tbl').rows[i].cells[10].innerHTML;	
		}
		self.opener.document.getElementById("total_money").value=parseFloat(self.opener.document.getElementById("total_money").value)+parseFloat(price.value)*parseFloat(number.value);
		self.opener.document.getElementById("zdprice").value=self.opener.document.getElementById("total_money").value;
		self.opener.document.getElementById("yfprice").value=self.opener.document.getElementById("total_money").value;
	} 
	else if ("<%response.write request.QueryString("type")%>"=="XD")
	{	
		if ("<%response.Write f_memoryprice%>" == "False")
		{
			price.value=document.getElementById('tbl').rows[i].cells[6].innerHTML;
		}else{
			price.value=document.getElementById('tbl').rows[i].cells[10].innerHTML;	
		}
		self.opener.document.getElementById("total_money").value=parseFloat(self.opener.document.getElementById("total_money").value)+parseFloat(price.value)*parseFloat(number.value);
	}
	else if ("<%response.write request.QueryString("type")%>"=="CD")
	{		
		if ("<%response.Write f_memoryprice%>" == "False")
		{
			price.value=document.getElementById('tbl').rows[i].cells[5].innerHTML;
		}else{
			price.value=document.getElementById('tbl').rows[i].cells[10].innerHTML;	
		}
		self.opener.document.getElementById("total_money").value=parseFloat(self.opener.document.getElementById("total_money").value)+parseFloat(price.value)*parseFloat(number.value);
	}
	else
	{
		if ("<%response.Write f_memoryprice%>" == "False")
		{
			price.value=document.getElementById('tbl').rows[i].cells[9].innerHTML;
		}else{
			price.value=document.getElementById('tbl').rows[i].cells[10].innerHTML;	
		}
		self.opener.document.getElementById("total_money").value=parseFloat(self.opener.document.getElementById("total_money").value)+parseFloat(price.value)*parseFloat(number.value);
	}
	money.value=price.value*number.value;	
	document.getElementById('tbl').rows[i].style.display="none";
}

function selectgoods(i){
var goodscode=self.opener.document.getElementById("totalgoodscode");
var goodsname=self.opener.document.getElementById("totalgoodsname");
var goodsunit=self.opener.document.getElementById("totalgoodsunit");
var units=self.opener.document.getElementById("totalunits");
var number=self.opener.document.getElementById("totalnumber");
goodscode.value=document.getElementById('tbl').rows[i].cells[1].innerText;
goodsname.value=document.getElementById('tbl').rows[i].cells[2].innerText;
goodsunit.value=document.getElementById('tbl').rows[i].cells[3].innerText;
units.value=document.getElementById('tbl').rows[i].cells[4].innerText;	 
number.value=1;
window.close();
}	

function selectgoodsproduct(i){
	var goodscode=self.opener.document.getElementById("totalgoodscode");
	var goodsname=self.opener.document.getElementById("totalgoodsname");
	var goodsunit=self.opener.document.getElementById("totalgoodsunit");
	var units=self.opener.document.getElementById("totalunits");
	var number=self.opener.document.getElementById("totalnumber");
	var price=self.opener.document.getElementById("totalprice");
	var aveprice=self.opener.document.getElementById("totalaveprice");
	goodscode.value=document.getElementById('tbl').rows[i].cells[1].innerText;
	goodsname.value=document.getElementById('tbl').rows[i].cells[2].innerText;
	goodsunit.value=document.getElementById('tbl').rows[i].cells[3].innerText;
	units.value=document.getElementById('tbl').rows[i].cells[4].innerText;	 
	number.value=1;
	price.value=document.getElementById('tbl').rows[i].cells[5].innerText;
	aveprice.value=document.getElementById('tbl').rows[i].cells[9].innerText;
	window.close();
}	
	
function selectcashcust(i){

	var custname=self.opener.document.getElementById("cust");
	
	
	//var mon=self.opener.document.getElementById("mon");
	custname.value=document.getElementById('tbl').rows[i].cells[2].innerHTML;
	
	//mon.value=document.getElementById('tbl').rows[i].cells[4].innerHTML;
	window.close();
}	

function selectinvoice(i){
	var bill=self.opener.document.getElementById("billcode");
	var cust=self.opener.document.getElementById("cust");
	var mon=self.opener.document.getElementById("money");
	bill.value=document.getElementById('tbl').rows[i].cells[1].innerHTML;
	cust.value=document.getElementById('tbl').rows[i].cells[2].innerHTML;
	mon.value=document.getElementById('tbl').rows[i].cells[6].innerHTML;
	window.close();
}	

function checkAll(e, itemName){
var aa = document.getElementsByName(itemName);
for (var j=0; j<aa.length; j++)
aa[j].checked = e.checked;
}
Object.prototype.clone = function() {
  var copy = (this instanceof Array) ? [] : {};
  for (attr in this) {
    if (!obj.hasOwnProperty(attr)) continue;
    copy[attr] = (typeof this[i] == "object")?obj[attr].clone():obj[attr];
  } 
  return copy;
};
</SCRIPT>

<%
Function SetSmallInt(DataValue)
	if (DataValue<1) and (DataValue>=0) then
	  if left(DataValue,1)<>"0" then
	    DataValue="0"&DataValue   
	  end if
	end if
	if (DataValue<0) and (DataValue>-1) then
	  if left(DataValue,2)<>"-0" then
	    DataValue = replace(DataValue,"-","-0")  
	  end if
	end if
	SetSmallInt = DataValue
End Function

function isnumber(str)
	if isnull(str) then
	isnumber=0
	else
	isnumber=cdbl(str)
	end if
end function

function showdepot1(s_name,depot_value)
  Response.Write "<select name=""" & s_name & """ id=""" & s_name & """ >"
  sql_depot = "select depotname from t_user where username='"&request.cookies("username")&"'"
  Set rs_depot = conn.Execute(sql_depot)
  
   arr = split(rs_depot("depotname"),",")
   
  if ubound(arr) <> -1 then
    for i = lbound(arr) to ubound(arr)-1
		  s_depotpower = s_depotpower & "'" & arr(i) & "',"
		  if depot_value = arr(i) then
		    Response.Write "<option value="&arr(i)&" selected>"&arr(i)&"</option>"
		  else
		    Response.Write "<option value="&arr(i)&">"&arr(i)&"</option>"
		  end if
		next
  end if
  Response.Write "</select>"
  rs_depot.close
  set rs_depot = nothing
end function

function showdepot(s_name,depot_value)

  Response.Write "<select name=""" & s_name & """ id=""" & s_name & """><option value=''>--请选择--</option>"
  sql_depot = "select depotname,RDepot from t_user where username='"&request.cookies("username")&"'"
 
set rs_depot = Server.CreateObject("adodb.recordset")
 rs_depot.open sql_depot, conn,1 ,1
 

if not rs_depot.eof then   
	
  if rs_depot("RDepot")=false then
  s_depotpower=""
  sql = "select depotname from t_depot where 1=1"
   set rs= Server.CreateObject("adodb.recordset")
 rs.open sql, conn,1 ,1
  
   Do While rs.eof=false
	  if depot_value = rs("depotname") then
	   Response.Write "<option value="&rs("depotname")&" selected>"&rs("depotname")&"</option>"
	  else
	   Response.Write "<option value="&rs("depotname")&">"&rs("depotname")&"</option>"
	  end if
	   rs.movenext
   loop 
 else
 	
  arr = split(rs_depot("depotname"),",")
  if ubound(arr) <> -1 then
    for i = lbound(arr) to ubound(arr)-1
      s_depotpower = s_depotpower & "'" & arr(i) & "',"
	next
  end if
  sql_depot = "select * from t_depot where depotname in ("&s_depotpower&"'a')"
  set rs_s_depot = Server.CreateObject("adodb.recordset")
 rs_s_depot.open sql_depot, conn,1 ,1
 
  Do While rs_s_depot.EOF = False
	  if depot_value = rs_s_depot("depotname") then
	    Response.Write "<option value="&rs_s_depot("depotname")&" selected>"&rs_s_depot("depotname")&"</option>"
	  else
	    Response.Write "<option value="&rs_s_depot("depotname")&">"&rs_s_depot("depotname")&"</option>"
	  end if
		rs_s_depot.movenext
  loop
  end if
  Response.Write "</select>"
  
  
  rs_s_depot.close
  set rs_s_depot = nothing
end if


end function

function showemployee(s_name,emp_name)
	Response.Write "<select name=""" & s_name & """ id=""" & s_name & """><option value=''>--请选择--</option>"
  sql_user = "select name from t_employee order by id"
  Set rs_user = conn.Execute(sql_user)
  Do While rs_user.EOF = False
	  if emp_name = rs_user("name") then
	    Response.Write "<option value="&rs_user("name")&" selected>"&rs_user("name")&"</option>"
	  else
	    Response.Write "<option value="&rs_user("name")&">"&rs_user("name")&"</option>"
	  end if
		rs_user.movenext
  loop
  Response.Write "</select>"
  rs_user.close
  set rs_user = nothing
end function

function showdepart(s_name,dep_name)
	Response.Write "<select name=""" & s_name & """ id=""" & s_name & """><option value=''>--请选择--</option>"
  sql_depart = "select name from t_department order by id"
  Set rs_depart = conn.Execute(sql_depart)
  Do While rs_depart.EOF = False
	  if dep_name = rs_depart("name") then
	    Response.Write "<option value="&rs_depart("name")&" selected>"&rs_depart("name")&"</option>"
	  else
	    Response.Write "<option value="&rs_depart("name")&">"&rs_depart("name")&"</option>"
	  end if
		rs_depart.movenext
  loop
  Response.Write "</select>"
  rs_depart.close
  set rs_depart = nothing
end function

function showdate(fieldname)
	Response.Write "<input type=""image"" src=""../img/date.gif"" height=""16"" width=""18"" style=""cursor:hand"" onClick=""calendar.show("&fieldname&");return false;"">"
end function

Private Function GetUrl() 
	Dim ScriptAddress,M_ItemUrl, M_item 
	ScriptAddress = CStr(Request.ServerVariables("SCRIPT_NAME")) '取得当前地址 
	M_ItemUrl = "" 
	If (Request.QueryString <> "") Then 
	ScriptAddress = ScriptAddress & "?" 
	For Each M_item In Request.QueryString 
	'如果页面传递参数是用page变量，那么判断一下page是否已经使用，避免重复！
	If InStr("page",M_Item)=0 Then
		if (InStr("order",M_Item)=0) and (InStr("orderkey",M_Item)=0) then
			M_ItemUrl = M_ItemUrl & M_Item &"="& Server.URLEncode(Request.QueryString(""&M_Item&"")) & "&"
		end if
	End If 
	Next 
	end if 
	GetUrl = ScriptAddress & M_ItemUrl
	
	if instr(1,GetUrl,"?",1)<1 then
	GetUrl=GetUrl&"?"
	
	end if
End Function 

'===========================
function showpage(sql,tablename,int_temp)
 dim starttime,endtime
 starttime = timer()
 
 sql_power = "select authority from t_user where username = '"& request.cookies("username") &"'"
 set rs_power = Server.CreateObject("adodb.recordset")
 rs_power.open sql_power, conn,1 ,1
 
 sql_sys = "select * from T_SoftInfo"
 set rs_sys = conn.Execute(sql_sys) '调用系统参数表
 
 Set mypage=new xdownpage '创建对象 
 mypage.getconn=conn '得到数据库连接 
 if Request.QueryString("order") = "" then
    if  tablename="showcash" then
	mypage.getsql= "select s_s_s.* from (" & sql & ") s_s_s order by adddate desc" & Request.QueryString("orderkey") 
	else
		mypage.getsql = sql
	end if
 
 else
 	mypage.getsql= "select s_s_s.* from (" & sql & ") s_s_s order by " & Request.QueryString("order") & " " & Request.QueryString("orderkey")
 end if
 
 mypage.pagesize=rs_sys("pagerecord")   '设置每一页的记录条数 
 set rs=mypage.getrs() '返回Recordset 
  
 Response.Write Request.Form("order")
 
 sql_field = "select * from t_fieldshow where t_fieldshow.show=1 and tablename='"&tablename&"' and hidden <> 1 order by showid"

 Set rs_field = server.CreateObject("adodb.recordset")
 'conn.cursorlocation=3
 rs_field.Open sql_field, conn, 1, 1
 

 sql_width = "select sum(width) as t_width from t_fieldshow where t_fieldshow.show=1 and hidden <> 1 and tablename='"&tablename&"'"
 Set rs_width = server.CreateObject("adodb.recordset")
 rs_width.Open sql_width, conn, 1, 1 



 sqlKey = "select field from t_fieldshow where tablename='"&tablename&"' and KeyWord = 1"
 Set rsKey = server.CreateObject("adodb.recordset")
 rsKey.Open sqlKey, conn, 1, 1 

 sqlColor = "select field from t_fieldshow where tablename='"&tablename&"' and Color = '1'"
 Set rsColor = server.CreateObject("adodb.recordset")
 rsColor.Open sqlColor, conn, 1, 1

 if (int_temp = 3) then
   Response.Write "<input type=""hidden"" name=""temp1"" id=""temp1"" onPropertyChange=""change_key()"" >"
 end if
 Response.Write "<input type=""hidden"" name=""temp"" id=""temp"">"'t_fieldshow中KeyWord字段在此显示
 Response.Write "</form>"
 Response.Write "<form name=""excel"" method=""post"" target=""_blank"" action=""../inc/excel.asp"">"
 Response.Write "<input type=""hidden"" name=""table"" value="""&tablename&""">"
 Response.Write "<input type=""hidden"" name=""sql"" value="""&server.HTMLEncode(sql)&""">"
 Response.Write "<table id=""tbl"" border=""0"" align=""left"" width='"&rs_width("t_width")&"px'>"
 Response.Write "<thead><tr align=""center"">"
 
 if Request.QueryString("orderkey") = "" then
 	sOrderKey = "desc"
 else
 	sOrderKey = ""
 end if
 if rsKey.recordcount = 0 then
 	Response.Write "<th width = '35'>序号</th>"
 else
 	Response.Write "<th width = '50'><input type='checkbox' id='checkall' onClick=""checkAll(this, 'del')"">选择</th>"
 end if
 if tablename="cashselectcust" then
  Response.Write "<th width = '35'>操作</th>"
 end if
 for k = 1 to rs_field.recordcount                 '显示表头
 	strWidth = rs_field("width")
 	if rs_field("showup")=1 and tablename<>"SaleTaxis"  then
		imgStr = ""
		if (Request.QueryString("order")=rs_field("field")) and (Request.QueryString("orderkey") = "desc") then
			imgStr = "<img border=0 src='../img/sanjiao1.gif'>"
		
		end if
		if (Request.QueryString("order")=rs_field("field")) and (Request.QueryString("orderkey") = "") then
			imgStr = "<img border=0 src='../img/sanjiao2.gif'>"
			
		end if
		
		Response.Write "<th width='"&strWidth&"px'><a href='#' class='a1' onclick=""document.form1.action='"&geturl()&"order="&rs_field("field")&"&orderkey="&sOrderkey&"';document.form1.submit();"">"&rs_field("name")&imgStr&"</a></th>"
	else
		Response.Write "<th width='"&strWidth&"px' >"&rs_field("name")&"</th>"	
	end if
  rs_field.movenext
 Next
 if tablename="SelectGoods" then
   Response.Write "<th width = '80'>选择</th>"
 end if
 Response.Write "</tr></thead><tbody>"
 if rs.recordcount = 0 then
 	Response.Write "<tr>"
	Response.Write "<td colspan="&rs_field.recordcount+1&"><div align=center><strong><font bold color=#ff0000><没有相关数据></font></strong></div></td>"
	Response.Write "</tr>"
 end if

 dim arr '定义合计数组
 for i =1 to 100
  redim arr(i)
 next
 for i = 1 to mypage.pagesize '接下来的操作就和操作一个普通Recordset对象一样操作 
 if not rs.eof then '这个标记是为了防止最后一页的溢出 
    if int_temp = 3 then
	  Response.Write "<tr id='tr"&i&"'  align=""center"" onMouseMove=""GetDIV('block',"&i&",'"&rs(""&rsKey("field")&"")&"')""  onmouseout=""GetDIV('none',"&i&",'"&rs(""&rsKey("field")&"")&"')"" onClick=""change()"">"
	elseif int_temp = 4 then
	  Response.Write "<tr id='tr"&i&"' align=""center"" onMouseOver=""showhint(this,'','','','"&rs("img")&"')"" onClick=""change()"" onMouseOut=""hidehintinfo()"">"
	else
      Response.Write "<tr id='tr"&i&"' align=""center"" onMouseOver=""over()"" onClick=""change()"" onMouseOut=""out()"">"
    end if
	if rsKey.recordcount = 0 then
		Response.Write "<th>"&i&"</th>"
	else
		Response.Write "<th><input type='checkbox' name='del' class='_del' value='"&rs(""&rsKey("field")&"")&"' id='id"&i&"'>"&i&"</th>"
	end if
	if tablename="cashselectcust" then
	Response.Write "<td><input type=""button""  id=""tb"" class=""button"" onclick =""selectcashcust("&i&")"" value='选择'/></td>"
	end if
	

 '这里以下的代码没有执行，需要调整
  rs_field.movefirst
  for k=1 to rs_field.recordcount
    str_show = rs(""&rs_field("field")&"")
	
	    if int_temp = 1 then
         Response.Write "<td id=""tb"">"'普通情况
  	    elseif int_temp = 2 then
	     Response.Write "<td id=""tb"" onDblClick=""CreateReturnValue(true,"&i&")"">"'选择货品
	    elseif int_temp = 3 then
	     Response.Write "<td id=""tb"">"
	    elseif int_temp = 4 then
	     Response.Write "<td id=""tb"">"'显示货品图片
		elseif (int_temp = 5) and ((k = 1) or (k = 2)) then
		 Response.Write "<th>"'显示为表头样式
		elseif int_temp = 6 then
		 Response.Write "<td id=""tb"" onDblClick=""selectgoods("&i&")"">"'模板中的选择货品
		elseif int_temp = 7 then
		 Response.Write "<td id=""tb"" onDblClick=""selectgoodsproduct("&i&")"">"'组装单中选择货品
		elseif int_temp = 8 then
		 Response.Write "<td id=""tb"" onDblClick=""selectcashcust("&i&")"">"'收付款单种的选择往来单位
		elseif int_temp = 9 then
		 Response.Write "<td id=""tb"" onDblClick=""selectinvoice("&i&")"">"'开票选择相关单据
		else
		 Response.Write "<td id=""tb"">"
	  end if

	if rs_field("DotSize") = 1 then '小数点位数
	  if isnumeric(str_show) then
	    str_show = (formatnumber(str_show,rs_sys("DotPrice"),-1,false,false))
	  else
	    str_show = (formatnumber(0,rs_sys("DotPrice"),-1,false,false))
	  end if
	elseif rs_field("DotSize") = 2 then
	  if isnumeric(str_show) then 
	    str_show = (formatnumber(str_show,rs_sys("DotNum"),-1,false,false))
	  else
	    str_show = (formatnumber(0,rs_sys("DotNum"),-1,false,false))
	  end if
	elseif rs_field("DotSize") = 3 then
	  if isnumeric(str_show) then 
	    str_show = (formatnumber(str_show,rs_sys("DotMon"),-1,false,false))
	  else
	    str_show = (formatnumber(0,rs_sys("DotMon"),-1,false,false))
	  end if
	else
	  str_show = str_show
	end if
	if rs_field("SumField") = 1 then
	  arr(k) = cdbl(arr(k)) + cdbl(str_show)
	  if isnumeric(str_show) then
	    arr(k) = (formatnumber(arr(k),rs_sys("DotPrice"),-1,false,false))
	  else
	    arr(k) = (formatnumber(0,rs_sys("DotPrice"),-1,false,false))
	  end if
	elseif rs_field("DotSize") = 2 then
	  if isnumeric(str_show) then 
	    arr(k) = (formatnumber(arr(k),rs_sys("DotNum"),-1,false,false))
	  else
	    arr(k) = (formatnumber(0,rs_sys("DotNum"),-1,false,false))
	  end if
	elseif rs_field("DotSize") = 3 then
	  if isnumeric(str_show) then 
	    arr(k) = (formatnumber(arr(k),rs_sys("DotMon"),-1,false,false))
	  else
	    arr(k) = (formatnumber(0,rs_sys("DotMon"),-1,false,false))
	  end if
	end if
  	if rsColor.recordcount > 0 then
		if isnumeric(rs(""&rsColor("field")&"")) then	
		  if cdbl(rs(""&rsColor("field")&"")) < 0 then
			  Response.Write str_show
			  ' Response.Write "<font color=#ff0000>"&str_show&"</font>"
		  else
			  Response.Write str_show
		  end if
		else
			Response.Write str_show
		end if
	else
		Response.Write str_show
	end if
				
	if int_temp = 5 and (k = 1 or k = 2) then
	  Response.Write "</th>"
	else
	  Response.Write "</td>"
	end if
	
    rs_field.movenext
	
  next
   if tablename="SelectGoods" then
 response.write "<td><img border=0 src='../images/chuliwb.png' style=cursor:pointer width=54 height=15 onDblClick=""CreateReturnValue(true,"&i&")""  onclick=""CreateReturnValue(true,"&i&")""></td>"
  end if
  Response.Write "</tr>"
 rs.movenext 
 else 
 exit for 
 end if 
 next
 Response.Write "</tbody>"
 sql_flag = "select sum(sumfield) as sum_flag from t_fieldshow where t_fieldshow.show=1 and t_fieldshow.tablename='"&tablename&"'"
 Set rs_flag = server.CreateObject("adodb.recordset")
 rs_flag.Open sql_flag, conn, 1, 1 
 if rs_flag("sum_flag") >= 1 then
   Response.Write "<tr onMouseOver=""over()"" onClick=""change()"" onMouseOut=""out()"">"
   Response.Write "<th>小计</th>"
   rs_field.movefirst
   for i = 1 to rs_field.recordcount
    Response.Write "<td align=""center"">"&arr(i)&"</td>"  
    rs_field.movenext()
   next
   Response.Write "</tr>"
 end if


 if rs_flag("sum_flag") >= 1 then
   Response.Write "<tr onMouseOver=""over()"" onClick=""change()"" onMouseOut=""out()"">"
   Response.Write "<th>合计</th>"
   rs_field.movefirst
   for i = 1 to rs_field.recordcount
   	if arr(i) <> "" then
			if i = 1 then
				strTemp = strTemp & "ifnull(sum("&rs_field("field")&"),0) as total"&rs_field("field")
			else
				strTemp = strTemp & ",ifnull(sum("&rs_field("field")&"),0) as total"&rs_field("field")
			end if
		else
			if i = 1 then
				if instr(sql,"@") <=0 then
					strTemp = strTemp & "'' as total"&rs_field("field")
				else
					strTemp = strTemp & "'''' as total"&rs_field("field")
				end if
			else
				if instr(sql,"@")<=0 then
					strTemp = strTemp & ",'' as total"&rs_field("field")
				else
					strTemp = strTemp & ",'''' as total"&rs_field("field")
				end if
			end if
		end if
    rs_field.movenext()
   next
 end if

 if rs_flag("sum_flag") >= 1 then
   i = instr(sql,"@")
   if i<=0 then
   	  sqlTotal = "select "&strTemp&" from ("&sql&") as sb"
   else
      sqla= left(sql,i-1)
	  sqlb=right(sql,len(sql)-i+1)
	  sqlTotal = Rtrim(sqla)&"Sum "&sqlb&",@field='"&strTemp&"'"
   end if 
   
   set rsTotal = Server.CreateObject("adodb.recordset")
   rsTotal.open sqlTotal,conn,1,1
 	rs_field.movefirst
	for i = 1 to rs_field.recordcount
		if rsTotal.recordcount = 0 then			
			Response.Write "<th align=center>" & "" & "</th>" 
		else
			Response.Write "<th align=center>" & rsTotal("total"&rs_field("field")) & "</th>" 
		end if
		rs_field.movenext
	next
	Response.Write "</tr>"
 end if
 if tablename="SelectGoods" then
  cp=1 
  else 
  cp=0 
 end if
 endtime = timer()-starttime
 Response.Write "<tr>"
 Response.Write "<th colspan="&rs_field.recordcount+1+cp&" align=""left"">"
 
 mypage.showpage()  '显示分页信息，这个方法可以，在set rs=mypage.getrs()以后任意位置调用，可以调用多次 
 
 Response.Write "<label class=""button1""><input type=""submit"" value=""Excel""></label>"
 '设置列显示功能
 Response.Write "<script>"
 Response.Write "function showtable(sTablename)"
 Response.Write "{"
 Response.Write "window.open ('../common/fieldshow.asp?tablename='+sTablename, 'newwindow', 'top=100,left=150,height=600, width=800,toolbar=no,menubar=no, scrollbars=yes, resizable=no,location=no, status=no')"
 Response.Write "}"
 Response.Write "</script>"
 if tablename <> "SelectGoods" then
 	Response.Write "<label class=""button1""><input type=""button"" onclick=showtable("""&tablename&""") value=""显示""></label>"
 '设置列显示功能结束
 end if
 response.write "搜索用时："&FormatNumber(endtime, 4, -1,false,false)&"(秒)"

 Response.Write "</th></tr>"
 Response.Write "</table>"
 Response.Write "</form>"
 
 rs_field.movefirst
 for i = 1 to rs_field.recordcount
 	if rs_field("pricepower") <> "" then
		if instr(rs_power("authority"),rs_field("pricepower")) = 0 then
			Response.Write "<script>setHiddenCol1(tbl,"&i&");</script>"
		end if
	end if
	rs_field.movenext
 next
 
	'释放数据集
	rs_sys.close
	set rs_sys = nothing
	rs_field.close
	set rs_field = nothing
	rs_width.close
	set rs_width = nothing
	rs_flag.close
	set rs_flag = nothing
end function

function ShowCombo(tablename,fieldname,comboname,combovalue)
	sql_combo = "select " & fieldname & " from " & tablename
	Set rs_combo = server.CreateObject("adodb.recordset")
	
	rs_combo.Open sql_combo, conn, 1, 1

	Response.Write "<select id="&comboname&" name="&comboname&">"
	if rs_combo.recordcount > 0 then
	rs_combo.movefirst
	Response.Write "<option value=''></option>"
	  for i = 1 to rs_combo.recordcount
		  if combovalue = rs_combo(fieldname) then
			  Response.Write "<option value=" & rs_combo(fieldname) & " selected>" & rs_combo(fieldname) & "</option>"
		  else
			  Response.Write "<option value=" & rs_combo(fieldname) & ">" & rs_combo(fieldname) & "</option>"
		  end if
		  rs_combo.movenext
	  next
	else
	Response.Write "<option value=''></option>"  
	end if
	Response.Write "</select>"
	rs_combo.close
	set rs_combo = nothing
end function

function formatdate(date)
	stryear=Year(Date)
	strMon=month(Date)
	strDay=day(date)
  If Len(strMon)<2 Then   
      strMon =   "0"   &   strMon   
  End   If   
  If   Len(strDay)<2   Then   
      strDay   =   "0"   &   strDay   
  End   If
  formatdate=stryear&"-"&strMon&"-"&strDay
end function

function showprice(fPrice)
	sqlSys = "select dotprice from t_softinfo"
	set rsSys = Server.CreateObject("adodb.recordset")
	rsSys.open sqlSys,conn,1,1
	showprice = formatnumber(fPrice,rsSys("dotprice"),-1,false,false)
end function

function shownumber(fNumber)
	sqlSys = "select dotNum from t_softinfo"
	set rsSys = Server.CreateObject("adodb.recordset")
	rsSys.open sqlSys,conn,1,1
	shownumber = formatnumber(fNumber,rsSys("dotnum"),-1,false,false)
end function

function showmoney(fAmt)
	sqlSys = "select dotmoney from t_softinfo"
	set rsSys = Server.CreateObject("adodb.recordset")
	rsSys.open sqlSys,conn,1,1
	showmoney = formatnumber(fAmt,rsSys("dotmoney"),-1,false,false)
end function


%>

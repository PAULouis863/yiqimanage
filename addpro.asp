<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理</title>
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->

<script language = "JavaScript">

var onecount;
onecount=0;
subcat = new Array();	
<%
i=0
sql="select * from [class] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof			''查询并且循环输出所有分类
	i=i+1	''设置数组下标，因此时ASP变量 i 的值为 1，而数组下标初始值为 0，所以下面的 i-1 就是为了符合数组下标的规则 
			''数据库中有多少符合的数据，其数组的总量就有多少，数组的最大值总是比数组的总量少 1，因为数组下标以 0 开头。 
%>
	subcat[<%=i-1%>] = new Array("<%=rs("mingcheng")%>","<%=rs("bigclassid")%>","<%=rs("id")%>");
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
onecount=<%response.Write i ''输出数组的总量，虽然ASP变量 i 在循环体外，但 i 在循环体内已经获得了最大值%>;

function changelocation(locationid)
{
	document.myform.classid.length = 0;
	var locationid=locationid;	
	var i;
	for (i=0;i < onecount; i++)		
	{
		if (subcat[i][1] == locationid)	
		{ 
			 document.myform.classid.options[document.myform.classid.length] = new Option(subcat[i][0], subcat[i][2]);
			 
		}        
	}
}    

</script>

<script language = "JavaScript">
function addpro()
{
	if(document.myform.jianjie.value=="") 
	{
		document.myform.jianjie.focus();
		alert("请输入仪器简介！");
		return false;
	}
	if(document.myform.riqi.value=="") 
	{
		document.myform.riqi.focus();
		alert("请输入添加日期！");
		return false;
	}
	if(document.myform.mingcheng.value=="") 
	{
		document.myform.mingcheng.focus();
		alert("请输入仪器名称！");
		return false;
	}
	if(document.myform.shichang.value=="") 
	{
		document.myform.shichang.focus();
		alert("请输入市场价格！");
		return false;
	}
	if(document.myform.huiyuan.value=="") 
	{
		document.myform.huiyuan.focus();
		alert("请输入会员价格！");
		return false;
	}
	if(document.myform.xinghao.value=="") 
	{
		document.myform.xinghao.focus();
		alert("请输入仪器型号！");
		return false;
	}
	if(document.myform.file.value=="") 
	{
		document.myform.file.focus();
		alert("请上传图片！");
		return false;
	}
	if(document.myform.shuoming.value=="") 
	{
		document.myform.shuoming.focus();
		alert("请输入仪器说明！");
		return false;
	}
	if(document.myform.beizhu.value=="") 
	{
		document.myform.beizhu.focus();
		alert("请输入仪器备注！");
		return false;
	}
	if(document.myform.shuliang.value=="") 
	{
		document.myform.shuliang.focus();
		alert("请输入仪器数量！");
		return false;
	}
}
</script>
</head>
<body>
<%
if request("action")="add" then	
	
	SafeRequest(trim(request("shuliang")))	
	SafeRequest(trim(request("shichang")))
	SafeRequest(trim(request("huiyuan")))
	sql="select * from [information]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("jianjie")=trim(request("jianjie"))
	rs("riqi")=request("riqi")
	rs("mingcheng")=request("mingcheng")
	rs("shichang")=request("shichang")
	rs("huiyuan")=request("huiyuan")
	rs("xinghao")=request("xinghao")
	rs("dengji")=request("dengji")
	rs("tupian")=request("file")
	rs("shuoming")=request("shuoming")
	rs("beizhu")=request("beizhu")
	rs("bigclassid")=request("bigclassid")
	rs("classid")=request("classid")
	rs("shuliang")=request("shuliang")
	rs("cishu")="1"
	rs.update
	rs.close
	set rs=nothing
	session("tupian")=""	
	response.Write("<script>alert('添加成功！');window.location.href='addpro.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
<form name="myform" action="addpro.asp" method="post" onSubmit="return addpro();">  <tr> 
    <td align="center"><font color="#FFFFFF">添加新的仪器</font></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="25%">仪器简介：</td>
          <td width="75%"><div align="left">
            <input name="jianjie" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>上架日期：</td>
          <td><div align="left">
            <input name="riqi" type="text" value="<%=now()%>" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>仪器名称：</td>
          <td><div align="left">
            <input name="mingcheng" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>出租价格：</td>
          <td><div align="left">
            <input name="shichang" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>购买价格：</td>
          <td><div align="left">
            <input name="huiyuan" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>仪器编号：</td>
          <td><div align="left">
            <input name="xinghao" type="text" size="30">
          </div></td>
        </tr>
		<tr bgcolor="#FFFFFF">
			<td rowspan="2" align="center">仪器图片：</td>
			<td><div align="left"><input id="file" type="text" name="file"/>
				<script>
				function backfn(fname){
					document.getElementById("file").value=fname;
				}
				</script></div>
			</td>
		</tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          
          <td><div align="left">
			<iframe width="300" height="24" scrolling="no" MARGINHEIGHT="0" MARGINWIDTH="0" name="upload" src="upload.asp"></iframe>
          </div></td>
        </tr>
		
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>仪器介绍：</td>
          <td><div align="left">
            <textarea cols="60" rows="8" name="shuoming"></textarea>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>仪器备注：</td>
          <td><div align="left">
            <textarea cols="60" rows="8" name="beizhu"></textarea>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td colspan="2">仪器等级：
            <select name="dengji">
                <option value="2">普通</option>
                <option value="1">精品</option>
            </select>
<%if request("bigclassid")="" and request("classid")="" then
%>
所属大类：

<select name="bigclassid" size="1" id="bigclassid" onChange="changelocation(document.myform.bigclassid.options[document.myform.bigclassid.selectedIndex].value)">

<%
sql="select * from [bigclass] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof	
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
所属小类：
<select name="classid">
<%
sql="select * from [class] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof	
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
<%else	%>
所属大类：
<select name="bigclassid" size="1" id="bigclassid">
<%
sql="select * from [bigclass] where id="&request("bigclassid")&" order by paixu;"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof	
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
所属小类：
<select name="classid">
<%
sql="select * from [class] where id="&request("classid")&" order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3	
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.close
set rs=nothing
%>
</select>
<%end if%>
&nbsp;&nbsp;仪器数量：<input name="shuliang" type="text" size="6">
<input name="action" type="hidden" value="add">
&nbsp;&nbsp;<input type="reset" name="reset" value="重写">
&nbsp;&nbsp;<input name="submit" type="submit" value="添加"></td>
          </tr>
      </table>
      <br></td>
  </tr>
</form>  
</table>
</body>
</html>
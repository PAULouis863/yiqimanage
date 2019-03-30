<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->
<script language = "JavaScript">
function addpro()
{
	if(document.myform.jianjie.value=="") 
	{
		document.myform.jianjie.focus();
		alert("请输入商品简介！");
		return false;
	}
	if(document.myform.shuliang.value=="") 
	{
		document.myform.shuliang.focus();
		alert("请输入商品数量！");
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
		alert("请输入商品名称！");
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
		alert("请输入商品型号！");
		return false;
	}
	if(document.myform.shuoming.value=="") 
	{
		document.myform.shuoming.focus();
		alert("请输入商品说明！");
		return false;
	}
	if(document.myform.beizhu.value=="") 
	{
		document.myform.beizhu.focus();
		alert("请输入商品备注！");
		return false;
	}
}
</script>
<%
if request("action")="update" then
	if request("jianjie")="" or request("riqi")="" or request("mingcheng")="" or request("shichang")="" or request("huiyuan")="" or request("dengji")="" or request("xinghao")="" or request("shuoming")="" or request("beizhu")="" or request("bigclassid")="" or request("classid")="" or request("shuliang")="" then
		response.Write("<script>alert('请详细填写！');history.back();</script>")
		response.End()
	end if
	SafeRequest(trim(request("shuliang")))
	SafeRequest(trim(request("shichang")))
	SafeRequest(trim(request("huiyuan")))
	sql="select * from [information] where id="&request("id")&""	'根据取得的商品 ID 进行修改
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs("jianjie")=trim(request("jianjie"))
	rs("riqi")=request("riqi")
	rs("shuliang")=request("shuliang")
	rs("mingcheng")=request("mingcheng")
	rs("shichang")=request("shichang")
	rs("huiyuan")=request("huiyuan")
	rs("xinghao")=request("xinghao")
	rs("dengji")=request("dengji")
	if session("tupian")<>"" then	'此处判断是否需要修改图片路径
		rs("tupian")=session("tupian")
	end if
	rs("shuoming")=request("shuoming")
	rs("beizhu")=request("beizhu")
	rs("bigclassid")=request("bigclassid")
	rs("classid")=request("classid")
	rs.update
	rs.close
	set rs=nothing
	session("tupian")=""
	response.Write("<script>alert('修改成功！');window.location.href='lookpro.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
<form name="myform" action="uppro.asp" method="post" onSubmit="return addpro();">  <tr> 
    <td align="center"><font color="#FFFFFF">修改影片</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
<%
sql="select * from [information] where id="&request("id")&";"	'查询对应的商品信息
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
%>
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="25%">影片简介：</td>
          <td width="75%"><div align="left">
            <input name="jianjie" type="text" value="<%=rs("jianjie")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>上架日期：</td>
          <td><div align="left">
            <input name="riqi" type="text" value="<%=rs("riqi")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>影片名称：</td>
          <td><div align="left">
            <input name="mingcheng" type="text" value="<%=rs("mingcheng")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>出租价格：</td>
          <td><div align="left">
            <input name="shichang" type="text" value="<%=rs("shichang")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>购买价格：</td>
          <td><div align="left">
            <input name="huiyuan" type="text" value="<%=rs("huiyuan")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>影片编号：</td>
          <td><div align="left">
            <input name="xinghao" type="text" value="<%=rs("xinghao")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>影片图片：</td>
          <td><div align="left">
            <iframe src="UpFile.asp" width="300" height="20" scrolling="no" MARGINHEIGHT="0" MARGINWIDTH="0" align="bottom"></iframe>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>详细说明：</td>
          <td><div align="left">
            <textarea cols="50" rows="6" name="shuoming"><%=rs("shuoming")%></textarea>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>影片备注：</td>
          <td><div align="left">
            <textarea cols="50" rows="6" name="beizhu"><%=rs("beizhu")%></textarea>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td colspan="2">
影片等级：
<select name="dengji">
<%if rs("dengji")="2" then	'判断是否为精品%>
<option value="2">普通</option>
<option value="1">精品</option>
<%else%>
<option value="1">精品</option>
<option value="2">普通</option>
<%end if%>
</select>
所属大类：
<select name="bigclassid">
<%
sql1="select * from [bigclass] where id="&rs("bigclassid")&";"	'查询该商品对应的大类
set rs1=Server.CreateObject("ADODB.Recordset")
rs1.open sql1,conn,3,3
%>
  <option value="<%=rs1("id")%>"><%=rs1("mingcheng")%></option>
<%
rs1.close
set rs1=nothing
%>
<%
sql2="select * from [bigclass] where id<>"&rs("bigclassid")&" order by paixu;"	'查询除了对应该商品的所有大类
'分两次查询的目的是既输入了所有类别，又做到了不重复输出
set rs2=Server.CreateObject("ADODB.Recordset")
rs2.open sql2,conn,3,3
do while not rs2.eof
%>
  <option value="<%=rs2("id")%>"><%=rs2("mingcheng")%></option>
<%
rs2.movenext
loop
rs2.close
set rs2=nothing
%>
</select>
所属小类：
<select name="classid">
<%
sql1="select * from [class] where id="&rs("classid")&";"	'查询该商品对应的小类
set rs1=Server.CreateObject("ADODB.Recordset")
rs1.open sql1,conn,3,3
%>
<option value="<%=rs1("id")%>"><%=rs1("mingcheng")%></option>
<%
rs1.close
set rs1=nothing
%>
<%
sql2="select * from [class] where id<>"&rs("classid")&" order by paixu;"	'查询除了对应该商品的所有小类，不重复输出
set rs2=Server.CreateObject("ADODB.Recordset")
rs2.open sql2,conn,3,3
do while not rs2.eof
%>
  <option value="<%=rs2("id")%>"><%=rs2("mingcheng")%></option>
<%
rs2.movenext
loop
rs2.close
set rs2=nothing
%>
</select>
&nbsp;&nbsp;影片数量：<input name="shuliang" type="text" size="6" value="<%=rs("shuliang")%>">
<input name="id" type="hidden" value="<%=request("id")%>">
<input name="action" type="hidden" value="update">
&nbsp;&nbsp;<input type="reset" name="reset" value="重写">
&nbsp;&nbsp;<input name="submit" type="submit" value="修改"></td>
          </tr>
<%
rs.close
set rs=nothing
%>
      </table>
      <br></td>
  </tr>
</form>  
</table>
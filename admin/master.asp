<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<!--#include file="../include/include.asp" -->
<script language = "JavaScript">
function addadmin()
{
	if(document.myform.user.value=="")
	{
		document.myform.user.focus();
		alert("请输入用户名！");
		return false;
	}
	if(document.myform.pass.value=="")
	{
		document.myform.pass.focus();
		alert("请输入密码！");
		return false;
	}
	if(document.myform.mail.value=="")
	{
		document.myform.mail.focus();
		alert("请输入电子邮件！");
		return false;
	}
	if(document.myform.xingming.value=="")
	{
		document.myform.xingming.focus();
		alert("请输入姓名！");
		return false;
	}
	if(document.myform.tel.value=="")
	{
		document.myform.tel.focus();
		alert("请输入电话！");
		return false;
	}
	if(document.myform.dizhi.value=="")
	{
		document.myform.dizhi.focus();
		alert("请输入地址！");
		return false;
	}
}
</script>
<%
if request("action")="add" then
	if trim(request("user"))="" or trim(request("pass"))="" or trim(request("mail"))="" or trim(request("xingming"))="" or trim(request("tel"))="" or trim(request("dizhi"))="" then
		response.Write("<script>alert('请详细填写！');history.back();</script>")
		response.End()
	end if
	sql="select * from [admin]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("username")=trim(request("user"))
	rs("pass")=md5(trim(request("pass")))
	rs("mail")=trim(request("mail"))
	rs("xingming")=trim(request("xingming"))
	rs("tel")=trim(request("tel"))
	rs("dizhi")=trim(request("dizhi"))
	rs("vip")="0"
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('添加成功！');window.location.href='master.asp';</script>")
end if
%>
<%
if request("action")="del" then
	conn.execute("delete from [admin]  where id="&request("id")&"")
	response.Write("<script>alert('删除成功！');window.location.href='master.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
  <tr> 
    <td align="center"><font color="#FFFFFF">查看后台管理员</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="133"><font color="#FF0000">用户I D</font></td>
          <td width="122"><font color="#FF0000">姓名</font></td>
          <td width="163"><font color="#FF0000">电子邮件</font></td>
          <td width="149"><font color="#FF0000">电话</font></td>
          <td width="233"><font color="#FF0000">地址</font></td>
          <td width="68"><font color="#FF0000">操 作</font></td>
        </tr>
<%
sql="select * from [admin] order by id desc;"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof And rs.bof then
	Response.Write "<p align='center' class='contents'> 对不起，暂无内容！</p>"
else
	rs.pagesize=10
	SafeRequest(request("page"))
	page=clng(request("page"))
	if page<1 then page=1
	if page>rs.pagecount then page=rs.pagecount
	show rs,page
	sub show(rs,page)
	rs.absolutepage=page
	for i=1 to rs.pagesize
%>
        <tr bgcolor="#FFFFFF" align="center"> 
            <td><a href="#" onClick='javascript:window.open("lookadmin.asp?id=<%=rs("id")%>");'><%=rs("username")%></a></td>
            <td><%=rs("xingming")%></td>
            <td><%=rs("mail")%></td>
            <td><%=rs("tel")%></td>
            <td><%=rs("dizhi")%></td>
            <td><a href="master.asp?action=del&id=<%=rs("id")%>">删除</a></td>
        </tr>
<%
	rs.movenext
	if rs.eof then exit for
	next
	end sub
%>      
      <tr>
	  <form action='' method='get' name='form'>
        <td colspan="6" bgcolor="#FFFFFF">
          <div align="center">
<%	
	if page<>1 then
		response.Write("&nbsp;&nbsp;<a href="&path&"?page=1>第一页</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&(page-1)&" >上一页</a>")
	end if 
	response.Write("&nbsp;&nbsp;当前 <font color='#FF0000'>"&page&"</font> 页")
	response.Write("&nbsp;&nbsp;条 <font color='#FF0000'>"&rs.recordcount&"</font>/<font color='#FF0000'>"&rs.pagecount&"</font> 页")
	if page<>rs.pagecount then
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&(page+1)&">下一页</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&rs.pagecount&">最末页</a>")
	end if
	response.Write("&nbsp;&nbsp;跳转到<input type='text' size='2' name='page'>页<input type='submit' value='GO'>")
end if
rs.close
set rs=nothing
%>
		&nbsp;<a href="outexcel4.asp" style="color:#F00">导出excel</a>    </div></td>
        </form></tr>
</table>
      <br></td>
  </tr>
</table>

<br>


<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
 <tr> 
    <td align="center"><font color="#FFFFFF">添加后台管理员</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
<form name="myform" action="master.asp" method="post" onSubmit="return addadmin();">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="25%">用户名：</td>
          <td width="75%"><div align="left">
            <input name="user" type="text">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>密&nbsp;&nbsp;码：</td>
          <td><div align="left">
            <input name="pass" type="password" size="20">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>E-MAIL：</td>
          <td><div align="left">
            <input name="mail" type="text">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>姓&nbsp;&nbsp;名：</td>
          <td><div align="left">
            <input name="xingming" type="text">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>电&nbsp;&nbsp;话：</td>
          <td><div align="left">
            <input name="tel" type="text">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>地&nbsp;&nbsp;址：</td>
          <td><div align="left">
            <input name="dizhi" type="text" size="40">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td colspan="2">
<input name="action" type="hidden" value="add">
<input name="submit" type="submit" value="添加">
&nbsp;&nbsp;<input type="reset" name="reset" value="重写">
&nbsp;&nbsp;</td>
          </tr>
</form>
      </table>
      <br></td>
  </tr>
</table>  
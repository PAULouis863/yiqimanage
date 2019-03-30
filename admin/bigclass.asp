<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->
<script language = "JavaScript">
function addbigclass()
{
	if(document.myform.mingcheng.value=="") 
	{
		document.myform.mingcheng.focus();
		alert("请输入大类名称！");
		return false;
	}
	if(document.myform.paixu.value=="") 
	{
		document.myform.paixu.focus();
		alert("请输入大类排序！");
		return false;
	}
}
</script>
<%
if request("action")="del" then
	conn.execute("delete from [bigclass]  where id="&request("id")&"")	'删除大类
	conn.execute("delete from [class]  where bigclassid="&request("id")&"")	'删除对应大类的分类
	conn.execute("delete from [information]  where bigclassid="&request("id")&"")	'删除对应大类的商品
	response.Write("<script>alert('删除成功！');window.location.href='bigclass.asp';</script>")
end if

if request("action")="update" then
	if request("paixu")="" or request("mingcheng")="" then
		response.Write("<script>alert('请详细填写！');history.back();</script>")
		response.End()
	end if
	SafeRequest(trim(request("paixu")))
	sql="select * from [bigclass] where id="&request("id")&""
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs("mingcheng")=trim(request("mingcheng"))
	rs("paixu")=trim(request("paixu"))
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('修改成功！');window.location.href='bigclass.asp';</script>")
end if

if request("action")="add" then
	if request("paixu")="" or request("mingcheng")="" then
		response.Write("<script>alert('请详细填写！');history.back();</script>")
		response.End()
	end if
	SafeRequest(trim(request("paixu")))
	sql="select * from [bigclass]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("mingcheng")=trim(request("mingcheng"))
	rs("paixu")=request("paixu")
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('添加成功！');window.location.href='bigclass.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
  <tr> 
    <td align="center"><font color="#FFFFFF">管理影片大类</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
		<tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="254"><font color="#FF0000">大类名称</font></td>
          <td width="152"><font color="#FF0000">大类排序</font></td>
          <td width="155"><font color="#FF0000">修改</font></td>
          <td width="150"><font color="#FF0000">操作</font></td>
          <td width="160"><font color="#FF0000">管理分类</font></td>
        </tr>
<%
sql="select * from [bigclass] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof
%>
        <form name="form" action="bigclass.asp" method="post">
        <tr bgcolor="#FFFFFF" align="center"> 
            <td><input name="mingcheng" type="text" value="<%=rs("mingcheng")%>" size="32"></td>
            <td><input name="paixu" type="text" size="10" value="<%=rs("paixu")%>"></td>
            <td><input name="submit" type="submit" value="修改"></td>
            <td><a href="bigclass.asp?action=del&id=<%=rs("id")%>">删除</a></td>
            <td><a href="class.asp?bigclassid=<%=rs("id")%>">管理分类</a></td>
        </tr>
		<input name="id" type="hidden" value="<%=rs("id")%>">
		<input name="action" type="hidden" value="update">
       </form> 
<%
rs.movenext
loop
%>
	</table>
      <br></td>
  </tr>
</table>
<br>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
  <tr>
    <td align="center"><font color="#FFFFFF">添加影片大类</font></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><br>
        <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <form name="myform" action="bigclass.asp" method="post" onSubmit="return addbigclass();">
          <tr height="20" bgcolor="#FFFFFF" align="center">
            <td width="425"><font color="#FF0000">大类名称&nbsp;&nbsp;              
              <input name="mingcheng" type="text" value="" size="32">
            </font></td>
            <td width="287"><font color="#FF0000">大类排序&nbsp;&nbsp;              
              <input name="paixu" type="text" size="10" value="">
            </font></td>
            <td width="165"><input name="submit" type="submit" value="添加"></td>
          </tr>
		  <input type="hidden" name="action" value="add">
		  </form>
        </table>
        <br></td>
  </tr>
</table>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->
<script language = "JavaScript">
function addclass()
{
	if(document.myform.mingcheng.value=="") 
	{
		document.myform.mingcheng.focus();
		alert("请输入分类名称！");
		return false;
	}
	if(document.myform.paixu.value=="") 
	{
		document.myform.paixu.focus();
		alert("请输入分类排序！");
		return false;
	}
}
</script>
<%
if request("action")="del" then
	conn.execute("delete from [class]  where id="&request("id")&"")
	conn.execute("delete from [information]  where classid="&request("id")&"")
	response.Write("<script>alert('删除成功！');window.location.href='class.asp?bigclassid="&request("bigclassid")&"';</script>")
end if

if request("action")="update" then
	if request("paixu")="" or request("mingcheng")="" then
		response.Write("<script>alert('请详细填写！');history.back();</script>")
		response.End()
	end if
	SafeRequest(trim(request("paixu")))
	sql="select * from [class] where id="&request("id")&""
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs("mingcheng")=trim(request("mingcheng"))
	rs("paixu")=trim(request("paixu"))
	if clng(rs("bigclassid"))<>clng(request("bigclassid")) then
		sql2="select * from [information] where bigclassid="&rs("bigclassid")&";"
		set rs2=Server.CreateObject("ADODB.Recordset")
		rs2.open sql2,conn,3,3
		do while not rs2.eof
			rs2("bigclassid")=request("bigclassid")
			rs2("classid")=rs("id")
			rs2.update
		rs2.movenext
		loop
		rs2.close
		set rs2=nothing
		rs("bigclassid")=request("bigclassid")
	end if
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('修改成功！');window.location.href='class.asp?bigclassid="&request("bigclassid")&"';</script>")
end if

if request("action")="add" then
	if request("paixu")="" or request("mingcheng")="" then
		response.Write("<script>alert('请详细填写！');history.back();</script>")
		response.End()
	end if
	SafeRequest(trim(request("paixu")))
	sql="select * from [class]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("mingcheng")=request("mingcheng")
	rs("paixu")=request("paixu")
	rs("bigclassid")=request("bigclassid")
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('添加成功！');window.location.href='class.asp?bigclassid="&request("bigclassid")&"';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
  <tr>
    <td align="center"><font color="#FFFFFF">管理影片分类</font></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><br>
        <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
          <tr height="20" bgcolor="#FFFFFF" align="center">
            <td width="254"><font color="#FF0000">分类名称</font></td>
            <td width="152"><font color="#FF0000">分类排序</font></td>
            <td width="155"><font color="#FF0000">所属大类</font></td>
            <td width="155"><font color="#FF0000">修改</font></td>
            <td width="150"><font color="#FF0000">操作</font></td>
            <td width="160"><font color="#FF0000">添加影片</font></td>
          </tr>
<%
if request("bigclassid")<>"" then	'判断是否指定大类
	bigclassid="where bigclassid="&request("bigclassid")&""	'增加查询大类的SQL语句赋值到变量 bigclassid 里
end if
sql="select * from [class] "&bigclassid&" order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof	'输出符合条件的分类
%>
<form name="form" action="class.asp" method="post">
            <tr bgcolor="#FFFFFF" align="center">
              <td><input name="mingcheng" type="text" value="<%=rs("mingcheng")%>" size="32"></td>
              <td><input name="paixu" type="text" size="10" value="<%=rs("paixu")%>"></td>
              <td>
<select name="bigclassid">
<%
sql2="select * from [bigclass] where id="&rs("bigclassid")&" order by paixu"
set rs2=Server.CreateObject("ADODB.Recordset")
rs2.open sql2,conn,3,3	'输出该分类所属的大类
%>
  <option value="<%=rs2("id")%>"><%=rs2("mingcheng")%></option>
<%
rs2.close
set rs2=nothing
%>
<%
sql3="select * from [bigclass] where id<>"&rs("bigclassid")&" order by paixu"
set rs3=Server.CreateObject("ADODB.Recordset")
rs3.open sql3,conn,3,3
do while not rs3.eof	'输出除了该分类所属的所有大类
%>
  <option value="<%=rs3("id")%>"><%=rs3("mingcheng")%></option>
<%
rs3.movenext
loop
rs3.close
set rs3=nothing
%>
</select>

			  </td>
              <td><input name="submit" type="submit" value="修改"></td>
              <td><a href="class.asp?action=del&id=<%=rs("id")%>&bigclassid=<%=rs("bigclassid")%>">删除</a></td>
              <td><a href="addpro.asp?bigclassid=<%=rs("bigclassid")%>&classid=<%=rs("id")%>">添加影片</a></td>
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
    <td align="center"><font color="#FFFFFF">添加影片分类</font></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><br>
        <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
          <form name="myform" action="class.asp" method="post" onSubmit="return addclass();">
            <tr height="20" bgcolor="#FFFFFF" align="center">
              <td width="285"><font color="#FF0000">分</font><font color="#FF0000">类名称&nbsp;&nbsp;
                    <input name="mingcheng" type="text" value="" size="26">
              </font></td>
              <td width="198"><font color="#FF0000">分类排序&nbsp;&nbsp;
                    <input name="paixu" type="text" size="10" value="">
              </font></td>
              <td width="304">
  <font color="#FF0000">所属大类;</font>&nbsp;&nbsp;
  <select name="bigclassid">
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
</select>			  </td>
              <td width="87"><input name="submit2" type="submit" value="添加"></td>
            </tr>
            <input type="hidden" name="action" value="add">
          </form>
        </table>
        <br></td>
  </tr>
</table>
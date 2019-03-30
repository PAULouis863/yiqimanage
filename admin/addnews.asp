<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<!--#include file="../include/include.asp" -->
<script language = "JavaScript">
function addnews()
{
	if(document.myform.user.value=="") 
	{
		document.myform.user.focus();
		alert("请输入发表人！");
		return false;
	}
	if(document.myform.shijian.value=="") 
	{
		document.myform.shijian.focus();
		alert("请输新闻时间！");
		return false;
	}
	if(document.myform.biaoti.value=="") 
	{
		document.myform.biaoti.focus();
		alert("请输新闻标题！");
		return false;
	}
	if(document.myform.neirong.value=="") 
	{
		document.myform.neirong.focus();
		alert("请输新闻内容！");
		return false;
	}
}
</script>
<%
if request("action")="add" then
	if trim(request("user"))="" or trim(request("shijian"))="" or trim(request("biaoti"))="" or trim(request("neirong"))="" then
		response.Write("<script>alert('请详细填写！');history.back();</script>")
		response.End()
	end if
	sql="select * from [news]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("user")=trim(request("user"))
	rs("shijian")=trim(request("shijian"))
	rs("biaoti")=trim(request("biaoti"))
	rs("neirong")=trim(request("neirong"))
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('站内新闻添加成功！');window.location.href='addnews.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
 <tr> 
    <td align="center"><font color="#FFFFFF">添加站内新闻</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
<form name="myform" action="addnews.asp" method="post" onSubmit="return addnews();">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="25%">发表人：</td>
          <td width="75%"><div align="left">
            <input name="user" type="text" value="<%=session("admin")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>时&nbsp;&nbsp;间：</td>
          <td><div align="left">
            <input name="shijian" type="text" value="<%=now()%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>标题名：</td>
          <td><div align="left">
            <input name="biaoti" type="text" size="40">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>内&nbsp;&nbsp;容：</td>
          <td><div align="left"><textarea name="neirong" cols="80" rows="18"></textarea>
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
      <br>
	  </td>
  </tr>
</table>
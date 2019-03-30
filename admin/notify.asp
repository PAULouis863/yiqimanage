<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<!--#include file="../include/include.asp" -->
<script language = "JavaScript">
function addnotify()
{
	if(document.myform.shijian.value=="") 
	{
		document.myform.shijian.focus();
		alert("请输入时间！");
		return false;
	}
	if(document.myform.neirong.value=="") 
	{
		document.myform.neirong.focus();
		alert("请输公告内容！");
		return false;
	}
}
</script>
<%
if request("action")="update" then
	if trim(request("shijian"))="" or trim(request("neirong"))="" then
		response.Write("<script>alert('请详细填写！');history.back();</script>")
		response.End()
	end if
	sql="select * from [placard]"	'在未指定操作条件时程序默认对表内的所有数据进行操作
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs("shijian")=trim(request("shijian"))
	rs("neirong")=trim(request("neirong"))
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('公告设置成功！');window.location.href='notify.asp';</script>")
end if
%>
<br>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
  <tr>
    <td align="center"><font color="#FFFFFF">站内公告设置</font></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><br>
        <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
<%
sql="select * from placard"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
'if not rs.eof then
%>
          <form name="myform" action="notify.asp" method="post" onSubmit="return addnotify();">
            <tr height="20" bgcolor="#FFFFFF" align="center">
              <td width="25%">时&nbsp;&nbsp;间：</td>
              <td width="75%"><div align="left">
                  <input name="shijian" type="text" value="<%=now()%>">
              </div></td>
            </tr>
            <tr height="20" bgcolor="#FFFFFF" align="center">
              <td>内&nbsp;&nbsp;容：</td>
              <td><div align="left">
                  <textarea name="neirong" cols="60" rows="8"><%=rs("neirong")%></textarea>
              </div></td>
            </tr>
            <tr height="20" bgcolor="#FFFFFF" align="center">
              <td colspan="2">
                <input name="action" type="hidden" value="update">
                <input name="submit" type="submit" value="设置">
&nbsp;&nbsp;
        <input type="reset" name="reset" value="重写">
&nbsp;&nbsp;</td>
            </tr>
          </form>
<%
'end if
rs.close
set rs=nothing
%>
        </table>
        <br></td>
  </tr>
</table>
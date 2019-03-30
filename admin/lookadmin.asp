<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->
<%
if request("action")="update" then
	sql="select * from [admin] where id="&request("id")&";"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	if rs("pass")<>trim(request("pass")) then
		rs("pass")=md5(trim(request("pass")))
		session("admin_pass")=md5(trim(request("pass")))
	end if
	rs("mail")=trim(request("mail"))
	rs("xingming")=trim(request("xingming"))
	rs("tel")=trim(request("tel"))
	rs("dizhi")=trim(request("dizhi"))
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('修改成功！');javascript:window.close();</script>")
end if
%>
<%
sql="select * from [admin] where id="&trim(request("id"))&";"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
<form name="myform" action="lookadmin.asp" method="post">  <tr> 
    <td align="center"><font color="#FFFFFF">查看/修改后台用户信息</font></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="60%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td width="25%">用户名：</td>
          <td width="75%"><div align="left"><%=rs("username")%>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>密&nbsp;&nbsp;码：</td>
          <td><div align="left">
              <input name="pass" type="password" size="33" value="<%=rs("pass")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>E-MAIL：</td>
          <td><div align="left">
              <input name="mail" type="text" size="30" value="<%=rs("mail")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>姓&nbsp;&nbsp;名：</td>
          <td><div align="left">
              <input name="xingming" type="text" size="30" value="<%=rs("xingming")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>电&nbsp;&nbsp;话：</td>
          <td><div align="left">
              <input name="tel" type="text" size="30" value="<%=rs("tel")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>地&nbsp;&nbsp;址：</td>
          <td><div align="left">
              <input name="dizhi" type="text" size="40" value="<%=rs("dizhi")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td colspan="2">
		  <input type="reset" name="reset" value="关闭" onClick='javascript:window.close();'>
            <input name="action" type="hidden" value="update">
            <input name="id" type="hidden" value="<%=request("id")%>">
&nbsp;&nbsp;
            <input name="submit" type="submit" value="修改">
&nbsp;&nbsp;</td>
        </tr>
      </table>
      <br></td>
  </tr>
</form>  
</table> 
<%
end if
rs.close
set rs=nothing
%>
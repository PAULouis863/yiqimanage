<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<!--#include file="../include/include.asp" -->
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
 <tr> 
    <td align="center"><font color="#FFFFFF">查看意见反馈</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
<%
sql="select * from [idea] where id="&request("id")&";"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
%>
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="25%">发表人：</td>
          <td width="75%"><div align="left"><%=rs("user")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>时&nbsp;&nbsp;间：</td>
          <td><div align="left"><%=rs("shijian")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>标题名：</td>
          <td><div align="left"><%=rs("biaoti")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="141">内&nbsp;&nbsp;容：</td>
          <td><div align="left"><%=HTMLEncode(rs("neirong"))%></div></td>
        </tr>
<%
rs.close
set rs=nothing
%>
      </table>
      <br></td>
  </tr>
</table>
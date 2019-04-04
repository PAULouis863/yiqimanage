<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
<%
sql="select * from [information] where id="&request("id")&""
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
%>
	<tr> 
    <td align="center"><font color="#FFFFFF">查看仪器</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td colspan="2" rowspan="7"><img src="../upfile/<%=rs("tupian")%>" alt="" width="320" height="300" border="0"></td>
          <td width="74" height="20"><font color="#FF0000">仪器名称</font></td>
          <td height="20" colspan="3"><%=rs("mingcheng")%></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td><font color="#FF0000">上架日期</font></td>
          <td><%=rs("riqi")%></td>
          <td><font color="#FF0000">仪器等级</font></td>
          <td><%if rs("dengji")="2" then response.Write("普通") else response.Write("精品") end if%></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td><font color="#FF0000">出租价格</font></td>
          <td width="111"><%=rs("shichang")%></td>
          <td width="62"><font color="#FF0000">购买价格格</font></td>
          <td width="122"><%=rs("huiyuan")%></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td><font color="#FF0000">仪器简介</font></td>
          <td colspan="3"><div align="left"><%=rs("jianjie")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td><font color="#FF0000">仪器编号</font></td>
          <td colspan="3"><div align="left"><%=rs("xinghao")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="105"><font color="#FF0000">详细说明</font></td>
          <td colspan="3"><%=HTMLEncode(rs("shuoming"))%></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="126"><font color="#FF0000">仪器备注</font></td>
          <td colspan="3"><%=HTMLEncode(rs("beizhu"))%></td>
        </tr>
      </table>
      <br></td>
  </tr>
<%
rs.close
set rs=nothing
%>
</table>
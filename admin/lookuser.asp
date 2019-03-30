<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->
<%  
if request("action")="update" then
	sql="select * from [user] where id="&request("id")&";"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	if rs("pass")<>trim(request("pass")) then
		rs("pass")=md5(trim(request("pass")))
	end if
	rs("mail")=trim(request("mail"))
	rs("youbian")=trim(request("youbian"))
	rs("xingming")=trim(request("xingming"))
	rs("shenfenzheng")=trim(request("shenfenzheng"))
	rs("tel")=trim(request("tel"))
	rs("qq")=trim(request("qq"))
	rs("tishi")=trim(request("tishi"))
	rs("huida")=trim(request("huida"))
	rs("dizhi")=trim(request("dizhi"))
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('修改成功！');javascript:window.close();</script>")
end if
%>
<%
sql="select * from [user] where id="&trim(request("id"))&";"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
<form name="myform" action="lookuser.asp" method="post">  <tr> 
    <td align="center"><font color="#FFFFFF">查看会员信息</font></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="60%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="19%">用&nbsp;户&nbsp;名：</td>
          <td width="81%"><div align="left"><%=rs("name")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>密&nbsp;&nbsp;&nbsp; 码：</td>
          <td><div align="left">
            <input name="pass" type="password" size="33" value="<%=rs("pass")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>电子邮件：</td>
          <td><div align="left">
            <input name="mail" type="text" value="<%=rs("mail")%>" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>邮&nbsp;&nbsp;&nbsp;&nbsp;编：</td>
          <td><div align="left">
            <input name="youbian" type="text" value="<%=rs("youbian")%>" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>姓&nbsp;&nbsp;&nbsp;&nbsp;名：</td>
          <td><div align="left">
            <input name="xingming" type="text" value="<%=rs("xingming")%>" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>电&nbsp;&nbsp;&nbsp;&nbsp;话：</td>
          <td><div align="left">
            <input name="tel" type="text" value="<%=rs("tel")%>" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>身&nbsp;份&nbsp;证：</td>
          <td><div align="left">
            <input name="shenfenzheng" type="text" value="<%=rs("shenfenzheng")%>" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>地&nbsp;&nbsp;&nbsp;&nbsp;址：</td>
          <td><div align="left">
            <input name="dizhi" type="text" value="<%=rs("dizhi")%>" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>联系&nbsp;&nbsp;QQ：</td>
          <td><div align="left">
            <input name="qq" type="text" value="<%=rs("qq")%>" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>提&nbsp;&nbsp;&nbsp;&nbsp;示：</td>
          <td><div align="left">
            <input name="tishi" type="text" value="<%=rs("tishi")%>" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>回&nbsp;&nbsp;&nbsp;&nbsp;答：</td>
          <td><div align="left">
            <input name="huida" type="password" value="<%=rs("huida")%>" size="33">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>注册时间：</td>
          <td><div align="left"><%=rs("shijian1")%>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>上次登录：</td>
          <td><div align="left"><%=rs("shijian2")%>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>登录次数：</td>
          <td><div align="left"><%=rs("cishu")%>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td colspan="2">
<input name="action" type="hidden" value="update">
<input name="id" type="hidden" value="<%=request("id")%>">
</td>
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
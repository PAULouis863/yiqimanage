<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<table width="197" height="753"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="197" valign="top" background="images/index_r2_c1.jpg"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="18">&nbsp;</td>
          </tr>
        </table>
<%
if session("user")<>"" then
	response.Write("<table width='100%'  border='0' cellspacing='0' cellpadding='0'><tr><td height='34'><div align='center'>您共登录了 "&session("cishu")+1&" 次<br></div></td></tr><tr><td height='34'><div align='center'>上次登录："&session("shijian")&" </div></td></tr><tr><td height='34'><div align='center'>欢迎 "&session("user")&" 用户</div></td></tr></table>")
else
%>
          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <form name="form" method="post" action="login.asp">
              <tr>
                <td width="14%" height="26">&nbsp;</td>
                <td width="29%"><span class="style1">用户名：</span></td>
                <td width="47%"><input name="user" type="text" size="12"></td>
                <td width="5%">&nbsp;</td>
                <td width="5%">&nbsp;</td>
              </tr>
              <tr>
                <td height="26">&nbsp;</td>
                <td><span class="style1">密&nbsp;&nbsp;码：</span></td>
                <td><input name="pass" type="password" size="12"></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td height="50">&nbsp;</td>
                <td><a href="reg.asp" target="_blank"><img src="images/reg.jpg" width="50" height="44" border="0"></a></td>
                <td><div align="right"><a href="#"><input type="image" src="images/login.jpg" border=0 name="images" width="79" height="44"></a></div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <input name="action" type="hidden" value="login">
            </form>
          </table>
<%end if%>		  
          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="76">&nbsp;</td>
            </tr>
          </table>
          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="163"><table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13%" height="161">&nbsp;</td>
                  <td width="79%"><table width="128"  border="0" align="center" cellpadding="0" cellspacing="0">
<%
sql="select * from [placard]"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then
%>
				    <tr>
                      <td width="128" height="150"><marquee width="128" height="150" direction="up" onmouseover="this.stop()" onmouseout="this.start()" scrollAmount="1"><%=rs("neirong")%></marquee></td>
                    </tr>
<%
end if
rs.close
set rs=nothing
%>
                  </table></td>
                  <td width="8%">&nbsp;</td>
                </tr>
              </table></td>
            </tr>
          </table>
          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="70">&nbsp;</td>
            </tr>
          </table>
          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="132"><table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13%" height="133">&nbsp;</td>
                  <td width="79%" valign="top"><table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="4"></td>
                    </tr>
<%
sql="select top 8 * from [news] order by id desc"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof
%>
                    <tr>
                      <td height="18"><a href="looknews.asp?id=<%=rs("id")%>" target="_blank">
<%
if len(trim(rs("biaoti")))>10 then
	response.write left(trim(rs("biaoti")),8)&"..."
else
	response.write trim(rs("biaoti"))
end if
%>					  
					  </a></td>
                    </tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
                  </table></td>
                  <td width="8%">&nbsp;</td>
                </tr>
              </table></td>
            </tr>
          </table></td>
      </tr>
    </table> 
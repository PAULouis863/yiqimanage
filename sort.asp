<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="include/conn.asp" -->
<!--#include file="include/include.asp" -->
<table width="793" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="3"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="792" height="165" background="images/index_r1_c1.jpg"><!--#include file="top.asp" --></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td width="197" height="753" valign="top" background="images/left2.jpg"><table width="168" height="39"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="147" height="50">&nbsp;</td>
        </tr>
      </table>
        <table width="168" height="700"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="147" height="232" valign="top">
<table cellspacing="0" cellpadding="0" width="158" align="center">
<%
i=0
sql="select * from [bigclass] order by paixu"	
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof
	i=i+1	
%>
                <tr>
                  <td onclick="showsubmenu(<%=i%>)" style="cursor:hand"  height="30">
                    <div align="center">
                      <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="18%">&nbsp;</td>
                          <td width="17%"><img src="images/tubiao.gif" width="23" height="17"></td>
                          <td width="65%" valign="bottom"><span><%response.Write rs("mingcheng")	%></span></td>
                        </tr>
                      </table>
                    </div></td></tr>
                <tr>
                  <td id="submenu<%=i%>" style="DISPLAY: none">
                    <div class="sec_menu" style="WIDTH: 158px">
                      <table cellpadding=0 cellspacing=0 align=center width=135>
<%
sql2="select * from [class] where bigclassid="&rs("id")&" order by id desc"	
set rs2=Server.CreateObject("ADODB.Recordset")
rs2.open sql2,conn,3,3
do while not rs2.eof
%>
                        <tr>
                          <td height=20><div align="center">
                            <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="50">&nbsp;</td>
                                <td width="83"><a href=class.asp?id=<%=rs2("id")%> target="class"><%response.Write rs2("mingcheng")	%></a></td>
                              </tr>
                            </table>
                          </div></td>
                        </tr>
<%
rs2.movenext
loop
%>
                      </table>
                    </div>
                    <br>
                  </td>
                </tr>
<%
rs.movenext
loop
%>
<script language=javascript>
function showsubmenu(sid)
{
	whichEl = eval("submenu" + sid);	
	if (whichEl.style.display == "none")	
	{
		eval("submenu" + sid + ".style.display=\"\";");	
	}
	else
	{
		eval("submenu" + sid + ".style.display=\"none\";");
	}
}
</script>
</table>
  </td>
          </tr>
      </table></td>
    <td width="590" valign="top"><table border="0" cellpadding="0" cellspacing="0" width="590">
      <tr>
        <td colspan="3"><img src="images/spfl.jpg" width="590" height="49"></td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r2_c1" src="images/index_7_r2_c1.jpg" width="590" height="9" border="0" alt=""></td>
      </tr>
      <tr>
        <td width="16" height="643" background="images/index_7_r3_c1.jpg">&nbsp;</td>
        <td width="565" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td>&nbsp;</td>
            </tr>
          </table>
            <table width="540" height="617"  border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td width="520" height="620" valign="top"><iframe name="class" src="class.asp" width="540" height="620" scrolling="no" MARGINHEIGHT="0" MARGINWIDTH="0" align="bottom" frameborder="0"></iframe></td>
              </tr>
          </table></td>
        <td width="10" background="images/index_7_r3_c3.jpg">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r4_c1" src="images/index_7_r4_c1.jpg" width="590" height="7" border="0" alt=""></td>
      </tr>
    </table></td>
    <td width="6"><img name="index_r2_c3" src="images/index_r2_c3.jpg" width="5" height="753" border="0" alt=""></td>
  </tr>
  <tr>
    <td colspan="3"><!--#include file="foot.asp" --></td>
  </tr>
</table> 
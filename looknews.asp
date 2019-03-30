<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>音像租赁管理系统</title>
<style type="text/css">
<!--
.style1 {color: #f2ab5b}
.style3 {font-weight: bold}
.style4 {color: #FF0000}
-->
</style>
<!--#include file="include/conn.asp" -->
<!--#include file="include/include.asp" -->
<table width="792" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="3"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="792" height="165" background="images/index_r1_c1.jpg"><!--#include file="top.asp" --></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td width="197"><!--#include file="left.asp" --></td>
    <td width="590" valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="590">
  <tr>
    <td colspan="3"><img name="index_7_r1_c1" src="images/znxw.jpg" width="590" height="49" border="0" alt=""></td>
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
      <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
<%
sql="select * from [news] where id="&request("id")&";"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
%>
        <tr>
          <td><div align="center"><span class="style4"><b><%=rs("biaoti")%></b></span></div></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td><%=HTMLEncode(rs("neirong"))%></td>
        </tr>
<%
rs.close
set rs=nothing
%>
      </table></td>
    <td width="10" background="images/index_7_r3_c3.jpg">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="3"><img name="index_7_r4_c1" src="images/index_7_r4_c1.jpg" width="590" height="7" border="0" alt=""></td>
  </tr>
</table></td>
    <td width="8"><img name="index_r2_c3" src="images/index_r2_c3.jpg" width="5" height="753" border="0" alt=""></td>
  </tr>
  <tr>
    <td colspan="3"><!--#include file="foot.asp" --></td>
  </tr>
</table>
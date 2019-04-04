<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<style type="text/css">
<!--
.style1 {color: #f2ab5b}
.style2 {
	color: #f37b54;
	font-weight: bold;
	font-size: 14pt;
}
.style3 {font-weight: bold}
.style4 {color: #000000}
-->
</style>
<!--#include file="include/conn.asp" -->
<!--#include file="include/include.asp" -->
<%
if request("action")="add" then
	if trim(request("comment"))="" then
		response.Write("<script>alert('请详细填写！');history.back();</script>")
		response.End()
	end if
	sql="select * from [comment]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("shangpinid")=request("Id")
	rs("shijian")=now()
	rs("comment")=request("comment")
	rs("mingcheng")=request("mingcheng")
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('评论发表成功！');window.location.href='lookpro.asp?id="&request("id")&"';</script>")
end if
%>
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
<%
sql="select * from [information] where id="&request("id")&""
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
%>
  <tr>
    <td colspan="3"><img name="index_7_r1_c1" src="images/spxx.jpg" width="590" height="49" border="0" alt=""></td>
  </tr>
  <tr>
    <td colspan="3"><img name="index_7_r2_c1" src="images/index_7_r2_c1.jpg" width="590" height="9" border="0" alt=""></td>
  </tr>
  <tr>
    <td width="16" height="643" background="images/index_7_r3_c1.jpg">&nbsp;</td>
    <td width="567" valign="top">
<table width="94%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="161" rowspan="7"><div align="center"><img src="upfile/<%=rs("tupian")%>" width="110" height="130" border="0"></div></td>
    <td width="22" height="16">&nbsp;</td>
    <td width="127"><span class="style4">【仪器名称】<%=rs("mingcheng")%></span></td>
    <td width="220" ><span class="style4">【仪器简介】<%=rs("jianjie")%></span></td>
  </tr>
  <tr>
    <td height="16">&nbsp;</td>
    <td><span class="style4">【出租价格】<%=rs("shichang")%></span></td>

    </tr>
  <tr>
    <td height="16">&nbsp;</td>
	<td><span class="style4"><a href="lookpro.asp?id=<%=rs("id")%>" target="_blank">【查看信息</a>】</span></td>
    
    <td><span class="style4">【上架日期】<%=rs("riqi")%></span></td>
  </tr>
  <tr>
    <td height="16">&nbsp;</td>></td>
    <td><span class="style4">【仪器编号】<%=rs("xinghao")%></span></td>
  </tr>
  <tr>
    <td height="19">&nbsp;</td>
    <td><span class="style4">【<a href="gouwu.asp?ProdId=<%=rs("id")%>">租借仪器</a>】</span></td>
    <td><span class="style4">【仪器等级】
      <%if rs("dengji")="2" then response.Write("精品") else response.Write("普通")%>
    </span></td>
  </tr>
  <tr>
    <td height="16">&nbsp;</td>
    <td><span class="style4">【浏览次数】<%=rs("cishu")%></span></td>
    <td><span class="style4">【仪器数量】<%=rs("shuliang")%></span></td>
  </tr>
</table>
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
    </table>
	<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="31%"><span class="style2">仪器介绍：</span></td>
        <td width="69%">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><%=HTMLEncode(rs("shuoming"))%></td>
      </tr>
    </table>
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
    </table>	
	<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="31%"><span class="style2">仪器备注：</span></td>
        <td width="69%">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><%=HTMLEncode(rs("beizhu"))%></td>
      </tr>
    </table>
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
    </table>
	<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td><span class="style2">用户评论：</span>&nbsp;&nbsp;&nbsp;&nbsp;<a href="lookcomment.asp?id=<%=rs("id")%>" target="_blank">查看用户评论</a></td>
      </tr>
      <tr>
        <td>
		</td>
      </tr>
    </table>
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
    </table>
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="172">
<form action="lookpro.asp" method="post">
  <div align="center">
  <textarea name="comment" cols="60" rows="12"></textarea>
  <input type="hidden" name="id" value="<%=rs("id")%>">
  <input type="submit" value="评论">
  <input type="hidden" name="action" value="add">
  <input type="hidden" name="mingcheng" value="<%=rs("mingcheng")%>">
  </div>
</form>		</td>
      </tr>
    </table></td>
    <td width="7" background="images/index_7_r3_c3.jpg">&nbsp;</td>
  </tr>
  <tr>
    <td height="7" colspan="3"><img name="index_7_r4_c1" src="images/index_7_r4_c1.jpg" width="590" height="7" border="0" alt=""></td>
  </tr>
<%
rs("cishu")=rs("cishu")+1
rs.update
rs.close
set rs=nothing
%>
</table></td>
    <td width="8"><img name="index_r2_c3" src="images/index_r2_c3.jpg" width="5" height="753" border="0" alt=""></td>
  </tr>
  <tr>
    <td colspan="3"><!--#include file="foot.asp" --></td>
  </tr>
</table>
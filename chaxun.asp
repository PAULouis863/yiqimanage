<!--#include file="include/conn.asp" -->
<!--#include file="include/include.asp" -->
<!--#include file="chk_user.asp" -->


<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.style1 {color: #f2ab5b}
-->
</style>
<%
if right(request("zhuangtai"),1)="3" then	
	sql="select * from [order] where didanhao='"&request("order")&"';"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs("zhuangtai")=right(request("zhuangtai"),1)	
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('用户确认成功！');window.location.href='chaxun.asp?order="&request("order")&"';</script>")
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
    <td width="590" valign="top"><table border="0" cellpadding="0" cellspacing="0" width="590">
      <tr>
        <td colspan="3"><img name="index_7_r1_c1" src="images/ddcx.jpg" width="590" height="49" border="0" alt=""></td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r2_c1" src="images/index_7_r2_c1.jpg" width="590" height="9" border="0" alt=""></td>
      </tr>
      <tr>
        <td width="16" height="643" background="images/index_7_r3_c1.jpg">&nbsp;</td>
        <td width="564" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table>
<%
if request("order")<>"" then
	SafeRequest(request("order"))
	sql="select * from [order] where didanhao='"&request("order")&"';"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	if not rs.eof then
		information=split(rs("information"),",")
		shuliang=split(rs("shuliang"),",")
%>

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>
      <table border="0" cellspacing="1" cellpadding="4" align="center" width="100%" bgcolor="BDBDBC">
	 <form name="form" action="chaxun.asp" method="post">
      <tr>
        <td width="13%" bgcolor="#FFFFFF">订单编号:</td>
        <td width="22%" bgcolor="#FFFFFF"><%=request("order")%></td>
        <td width="18%" bgcolor="#FFFFFF"><div align="center">已收款
              <input type="checkbox" name="zhuangtai" disabled value="1"<%if rs("zhuangtai")>0 then response.Write("checked") end if	%>>
        </div></td>
        <td width="18%" bgcolor="#FFFFFF"><div align="center">已发货
              <input type="checkbox" name="zhuangtai" disabled value="2"<%if rs("zhuangtai")>1 then response.Write("checked") end if	%>>
        </div></td>
        <td width="18%" bgcolor="#FFFFFF"><div align="center">已收货
              <input type="checkbox" name="zhuangtai" <%if rs("zhuangtai")<2 or rs("zhuangtai")>2 then response.Write("disabled") end if%> value="3"<%if rs("zhuangtai")>2 then response.Write("checked") end if %>>
        </div></td>
        <td width="18%" bgcolor="#FFFFFF"><input type="submit" value="确认收货"></td>
      </tr>
	  <input type="hidden" name="order" value="<%response.Write request("order") %>">
	  </form>
    </table>
      <table border="0" cellspacing="1" cellpadding="4" align="center" width="100％" bgcolor="BDBDBC">
        <tr bgcolor="#FFFFFF" height="25" align="center">
          <td width="300">商 品 名 称</td>
          <td width="40">数量</td>
          <td width="60">出租价格</td>
          
          <td width="60">成交价</td>
          <td width="70">小 计</td>
        </tr>
        <tr bgcolor="#FFFFFF" height="25"align="center">
          <td align="left">
<%
		for i=0 to ubound(information)
			sql2="select * from [information] where id="&trim(information(i))&""
			set rs2=Server.CreateObject("ADODB.Recordset")
			rs2.open sql2,conn,3,3
%>
<a href="lookpro.asp?id=<%=rs2("id")%>" target="_blank"><%=rs2("mingcheng")%></a><br>
<%
			rs2.close
			set rs2=nothing
		next
%>
		  </td>
          <td>
<%
		for i=0 to ubound(shuliang)
			response.Write(shuliang(i))
			response.Write("<br>")
		next
%>
		</td>
          <td>
<%
		for i=0 to ubound(shuliang)
			sql2="select * from [information] where id="&trim(information(i))&""
			set rs2=Server.CreateObject("ADODB.Recordset")
			rs2.open sql2,conn,3,3
			response.Write(rs2("shichang"))
			response.Write("<br>")
			rs2.close
			set rs2=nothing
		next
%>
		  </td>
          <td>
<%
		for i=0 to ubound(shuliang)
			sql2="select * from [information] where id="&trim(information(i))&""
			set rs2=Server.CreateObject("ADODB.Recordset")
			rs2.open sql2,conn,3,3
			response.Write(rs2("huiyuan"))
			response.Write("<br>")
			rs2.close
			set rs2=nothing
		next
%>
		  </td>
          <td>
<%
		for i=0 to ubound(shuliang)
			sql2="select * from [information] where id="&trim(information(i))&""
			set rs2=Server.CreateObject("ADODB.Recordset")
			rs2.open sql2,conn,3,3
			response.Write(rs2("huiyuan"))
			response.Write("<br>")
			rs2.close
			set rs2=nothing
		next
%>
		  </td>
          <td>
<%
		for i=0 to ubound(shuliang)
			sql2="select * from [information] where id="&trim(information(i))&""
			set rs2=Server.CreateObject("ADODB.Recordset")
			rs2.open sql2,conn,3,3
			response.Write(rs2("huiyuan")*shuliang(i))
			response.Write("<br>")
			rs2.close
			set rs2=nothing
		next
%>
		  </td>
        </tr> 
    </table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="BDBDBC">
        <tr bgcolor=#ffffff>
          <td width="150" height="26"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="10"></td>
            </tr>
          </table>
            收货人姓名：
            <table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="10"></td>
              </tr>
            </table></td>
          <td width="600" height="26"><%=rs("shoujianren")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td height="30">详细地址：</td>
          <td height="28"><%=rs("dizhi")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td height="30">邮　　编：</td>
          <td height="28"><%=rs("youbian")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td height="30">电　　话：</td>
          <td height="28"><%=rs("tel")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td height="30">电子邮件：</td>
          <td height="28"><%=rs("mail")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td height="26">送货方式：</td>
          <td height="28"><%=rs("songhuo")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td height="30">支付方式：</td>
          <td height="20"><%=rs("zhifu")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td height="40" valign="middle">简单留言：</td>
          <td height="28"><%=HTMLEncode(rs("leaveword"))%></td>
        </tr>
    </table>	</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" class="table-zuoyou">
  <tr>
    <td> <br></td>
  </tr>
  <tr>
    <td>
      </td>
  </tr>
</table>
<%
	else
		response.Write("<script>alert('无此订单号');window.location.href='chaxun.asp';</script>")
	end if
	rs.close
	set rs=nothing
else
%>
<form name="form" action="chaxun.asp" method="post">
  <div align="center">请输入您的订单号：&nbsp;&nbsp;
      <input name="order" type="text">
      <input type="submit" value="查询">
  </div>
</form>
<%end if%>
          <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td>&nbsp;</td>
            </tr>
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
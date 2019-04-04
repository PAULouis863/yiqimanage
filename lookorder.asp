<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<!--#include file="../include/include.asp" -->
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
<%  Sum=0
if request("order")<>"" and request("action")="update" then	'判断是否修改
	sql="select * from [order] where didanhao='"&request("order")&"';"	'按订单号查询
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	if not rs.eof then
		if right(request("zhuangtai"),1)>="2" and rs("zhuangtai")<"2" then
		'此时 request("zhuangtai") 为字符串形式，例：1,2或1,2,3 这个时候我们只要最右边的字符就可以知道提交的值了
		'当 right(request("zhuangtai"),1) 的值大与或等于 2 （已经发货了）并且数据库中的值小于 2 (仪器数量没有被修改过)的时候才对仪器的数量进行修改
			information=split(rs("information"),",")
			shuliang=split(rs("shuliang"),",")
			for i=0 to ubound(information)	'循环输出仪器 ID，有多少仪器就对应仪器 ID进行数量修改
				sql2="select * from [information] where id="&information(i)&""
				set rs2=Server.CreateObject("ADODB.Recordset")
				rs2.open sql2,conn,3,3
				rs2("shuliang")=rs2("shuliang")-shuliang(i)	'新的仪器数量=原仪器数量-订单中的仪器数量
				rs2.update
				rs2.close
				set rs2=nothing
			next
		end if
		if request("zhuangtai")<>"" then
			rs("zhuangtai")=right(request("zhuangtai"),1)
		else
			rs("zhuangtai")=0	'如果订单状态不属于已收款、已发货、已收货的情况下
		end if
		rs.update
		rs.close
		set rs=nothing
		response.Write("<script>alert('修改订单成功！');window.location.href='lookorder.asp?order="&request("order")&"';</script>")
	end if
end if
if request("order")<>"" then
	SafeRequest(request("order"))	'判断订单号是否为数字型
	sql="select * from [order] where didanhao='"&request("order")&"';"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	if not rs.eof then
		information=split(rs("information"),",")	'用来分割字符串，首先要拆分此订单中每个仪器的 ID
		shuliang=split(rs("shuliang"),",")	'仪器对应的购买数量
        session("je")=rs("xf")
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
    <tr>
      <td align="center"><font color="#FFFFFF">仪器订单管理</font></td>
    </tr>
    <tr>
      <td valign="top" bgcolor="#FFFFFF"><br>
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td><table border="0" cellspacing="1" cellpadding="4" align="center" width="100％" bgcolor="#6699FF">
	<form action="lookorder.asp" method="get">
      <tr>
        <td width="13%" bgcolor="#FFFFFF">订单编号:</td>
        <td width="22%" bgcolor="#FFFFFF"><%=rs("didanhao")%></td>
        <td width="18%" bgcolor="#FFFFFF"><div align="center">已收款
              <input type="checkbox" name="zhuangtai" value="1"<%if rs("zhuangtai")>0 then response.Write("checked") end if	'确定仪器状态%>>
        </div></td>
        <td width="18%" bgcolor="#FFFFFF"><div align="center">已发货
              <input type="checkbox" name="zhuangtai" value="2"<%if rs("zhuangtai")>1 then response.Write("checked") end if%>>
        </div></td>
        <td width="18%" bgcolor="#FFFFFF"><div align="center">已收货
              <input type="checkbox" name="zhuangtai" value="3"<%if rs("zhuangtai")>2 then response.Write("checked") end if%>>
        </div></td>
        <td width="7%" bgcolor="#FFFFFF"><div align="right">
          <input type="submit" name="submit" value="修改">
        </div></td>
      </tr>
	  <input type="hidden" name="order" value="<%response.Write rs("didanhao")	'将订单号作为隐藏值进行提交%>">
	  <input type="hidden" name="action" value="update">
	  </form>
    </table></td>
  </tr>
  <tr>
    <td>
      <table border="0" cellspacing="1" cellpadding="4" align="center" width="100％" bgcolor="#6699FF">
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
		for i=0 to ubound(information)	'此时变量 information 为数组形式
			sql2="select * from [information] where id="&trim(information(i))&""	'得到每个仪器的 ID
			set rs2=Server.CreateObject("ADODB.Recordset")
			rs2.open sql2,conn,3,3
%>
<a href="pro.asp?id=<%=rs2("id")%>"><%response.Write rs2("mingcheng") '输出仪器名称%></a><br>
<%
			rs2.close
			set rs2=nothing
		next
%>
		  </td>
          <td>
<%
		for i=0 to ubound(shuliang)
			response.Write(shuliang(i)) '输出仪器数量
			response.Write("<br>")
		next
%>
		</td>
          <td>
<%
		for i=0 to ubound(shuliang)	'此时数组 shuliang 和 information 的最大下标是一样的
			'因为在存储订单时仪器对应着数量，而这里的 FOR 循环语句需要最大值，所以任何一个都可以作为 FOR 循环的最大值
			sql2="select * from [information] where id="&trim(information(i))&""
			set rs2=Server.CreateObject("ADODB.Recordset")
			rs2.open sql2,conn,3,3
			response.Write(rs2("shichang")) '输出仪器市场价格
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
			response.Write(rs2("shichang")) '输出仪器会员价格
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
                 Sum=Sum+rs2("shichang")*shuliang(i)
			response.Write(rs2("shichang")*shuliang(i))
			response.Write("<br>")
			rs2.close
			set rs2=nothing
		next
%>
		  </td>
        </tr>        <tr bgcolor="#FFFFFF" height="25"align="center">
          <td align="left" colspan="5">
            <%if session("je")<100 Then response.Write("铜牌会员九折 总计："&Sum*0.9) else if  session("je")>100 and session("je")<500 Then response.Write("银牌会员八折 总计："&Sum*0.8)   else if   session("je")>500 Then response.Write("金牌会员七折 总计："&Sum*0.7)    end if %></td>
        </tr> 
    </table></td>
  </tr>
</table>
<table width="600"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="center" class="style1">注：仪器确定发货后，该仪器数量将自动从库存中相应减少！</div></td>
  </tr>
</table>
<table width="600" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6699FF">
        <tr bgcolor=#ffffff>
          <td width="150">收货人姓名：</td>
          <td width="600" height="28"><%=rs("shoujianren")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td>详细地址：</td>
          <td height="28"><%=rs("dizhi")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td>邮　　编：</td>
          <td height="28"><%=rs("youbian")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td>电　　话：</td>
          <td height="28"><%=rs("tel")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td>电子邮件：</td>
          <td height="28"><%=rs("mail")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td>送货方式：</td>
          <td height="28"><%=rs("songhuo")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td>支付方式：</td>
          <td height="20"><%=rs("zhifu")%></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td>简单留言：</td>
          <td height="28"><%response.Write HTMLEncode(rs("leaveword"))	'函数 HTMLEncode 的功能就是替换空格、换行,代码在 include.asp 文件里%></td>
        </tr>
</table>
<%
	else
		response.Write("<script>alert('无此订单号');</script>")
	end if
	rs.close
	set rs=nothing
end if
%>		  <br></td>
    </tr>
</table>

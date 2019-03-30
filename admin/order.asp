<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<!--#include file="../include/include.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>管理用户订单</title>
</head>
<body>
<%
if request("action")="del" then	
	conn.execute("delete from [order]  where id="&request("id")&"")	
	response.Write("<script>alert('删除成功！');window.location.href='order.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
  <tr> 
    <td align="center"><font color="#FFFFFF">管理用户订单</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="151"><font color="#FF0000">订单编号</font></td>
          <td width="147"><font color="#FF0000">收货方式</font></td>
          <td width="168"><font color="#FF0000">付款方式</font></td>
          <td width="156"><font color="#FF0000">下单用户</font></td>
          <td width="158"><font color="#FF0000">下单时间</font></td>
          <td width="88"><font color="#FF0000">操 作</font></td>
        </tr>
<%
if request("tiaojian")="1" then
	chaxun="where didanhao='"&request("guanjian")&"'"	
end if
if request("tiaojian")="2" then
	chaxun="where name='"&request("guanjian")&"'"	
end if
if request("tiaojian")="3" then
	if request("guanjian")="" then	
		chaxun="where zhuangtai='1'"	
	else
		chaxun="where zhuangtai='1' and  name='"&request("guanjian")&"'"	
	end if
end if
if request("tiaojian")="4" then
	if request("guanjian")="" then	
		chaxun="where zhuangtai='2'"	
	else
		chaxun="where zhuangtai='2' and  name='"&request("guanjian")&"'"	'按已发货的订单并且对应用户名查询
	end if
end if
sql="select * from [order] "&chaxun&" order by id desc;"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof And rs.bof then	
	Response.Write "<p align='center' class='contents'> 对不起，暂无内容！</p>"
else
	rs.pagesize=10	
	SafeRequest(request("page"))
	page=clng(request("page"))
	if page<1 then page=1	
	if page>rs.pagecount then page=rs.pagecount	
	show rs,page
	sub show(rs,page)	
	rs.absolutepage=page
	for i=1 to rs.pagesize
%>
        <tr bgcolor="#FFFFFF" align="center"> 
            <td><a href="lookorder.asp?order=<%=rs("didanhao")%>"><%=rs("didanhao")%></a></td>
            <td><%=rs("songhuo")%></td>
            <td><%=rs("zhifu")%></td>
            <td><%=rs("name")%></td>
            <td><%=rs("shijian")%></td>
            <td><a href="order.asp?action=del&id=<%=rs("id")%>">删除</a></td>
        </tr>
<%
	rs.movenext
	if rs.eof then exit for	
	next
	end sub	
%>      
        <tr bgcolor="#FFFFFF" align="center">
		<form name="form" action="?" method="get">		
          <td colspan="6">
<%	
	if page<>1 then
		response.Write("&nbsp;&nbsp;<a href="&path&"?page=1>第一页</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&(page-1)&" >上一页</a>")
	end if 
	response.Write("&nbsp;&nbsp;当前 <font color='#FF0000'>"&page&"</font> 页")
	response.Write("&nbsp;&nbsp;条 <font color='#FF0000'>"&rs.recordcount&"</font>/<font color='#FF0000'>"&rs.pagecount&"</font> 页")
	if page<>rs.pagecount then
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&(page+1)&">下一页</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&rs.pagecount&">最末页</a>")
	end if
	response.Write("&nbsp;&nbsp;跳转到<input type='text' size='2' name='page'>页<input type='submit' value='GO'>")
end if
rs.close
set rs=nothing
%>
		  </td>
		  </form>
        </tr>
</table>
      <br></td>
  </tr>
</table>
<br>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
<form action="order.asp" method="post">
  <tr>
    <td align="center"><font color="#FFFFFF">用户订单查询</font></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><div align="center">查询条件&nbsp;&nbsp;
        <input name="guanjian" type="text" size="30">
        &nbsp;&nbsp;
	      <select name="tiaojian">
            <option value="1">订单编号查询</option>
            <option value="2">用户订单查询</option>
            <option value="3">付款订单查询</option>
            <option value="4">发货订单查询</option>
          </select>
    &nbsp;&nbsp;<input type="submit" value="查询"></div></td>
  </tr>
  <input type="hidden" name="chaxun" value="chaxun">
 </form>
</table>
</body>
</html>
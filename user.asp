<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<!--#include file="../include/include.asp" -->
<%
if request("action")="del" then
	conn.execute("delete from [user]  where id="&request("id")&"")
	response.Write("<script>alert('删除成功！');window.location.href='user.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
  <tr> 
    <td align="center"><font color="#FFFFFF">用户信息管理</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="133"><font color="#FF0000">用户I D</font></td>
          <td width="122"><font color="#FF0000">姓名</font></td>
          <td width="163"><font color="#FF0000">电子邮件</font></td>
          <td width="149"><font color="#FF0000">电话</font></td>
            <td width="149"><font color="#FF0000">会员等级</font></td>
          <td width="233"><font color="#FF0000">地址</font></td>
          <td width="68"><font color="#FF0000">操 作</font></td>
        </tr>
<%
sql="select * from [user] order by id desc;"
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
            <td><a href="#" onClick='javascript:window.open("lookuser.asp?id=<%=rs("id")%>");'><%=rs("name")%></a></td>
            <td><%=rs("xingming")%></td>
            <td><%=rs("mail")%></td>
            <td><%=rs("tel")%></td>
             <td><%if rs("xf")<100 Then response.Write("铜牌会员") else if   rs("xf")>100 and rs("xf")<500 Then response.Write("银牌会员")   else if    rs("xf")>500 Then response.Write("金牌会员")    end if %></td>
            <td><%=rs("dizhi")%></td>
            <td><a href="user.asp?action=del&id=<%=rs("id")%>">删除</a></td>
        </tr>
<%
	rs.movenext
	if rs.eof then exit for
	next
	end sub
%>
      <tr>
	  <form action='' method='get' name='form'>
        <td colspan="6" bgcolor="#FFFFFF">
          <div align="center">
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
%>&nbsp;<a href="outexcel3.asp" style="color:#F00">导出excel</a>
		    </div></td>
        </form></tr>
</table>
      <br></td>
  </tr>  
</table>
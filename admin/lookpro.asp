<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->
<%
if request("action")="del" then
	conn.execute("delete from [information]  where id="&request("id")&"")
	response.Write("<script>alert('删除成功！');window.location.href='lookpro.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
	<tr> 
    <td align="center"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
<form name="myform" action="lookpro.asp" method="post">  <tr> 
        <tr>
          <td width="320">&nbsp;</td>
          <td width="120"><div align="center"><font color="#FFFFFF">影片信息管理</font> </div></td>
          <td width="160">
<font color="#FFFFFF">所属大类：</font>
<select name="bigclassid">
<%
sql="select * from [bigclass] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select></td>
          <td width="160">
<font color="#FFFFFF">所属小类：</font>
<select name="classid">
<%
sql="select * from [class] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select></td>
          <td width="40"><input type="submit" value="查看"></td>
        </tr>
	</form>
      </table></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="234"><font color="#FF0000">影片名称</font></td>
          <td width="167"><div align="center"><font color="#FF0000">上架时间</font></div></td>
          <td width="139"><font color="#FF0000">出租价格</font></td>
          <td width="132"><font color="#FF0000">购买价格格</font></td>
          <td width="100"><font color="#FF0000">修改</font></td>
          <td width="96"><font color="#FF0000">操作</font></td>
        </tr>
<%
if request("bigclassid")<>"" and request("classid")<>"" then	'判断是否按指定类别进行查询
	wh="where bigclassid="&request("bigclassid")&" and classid="&request("classid")&""
end if
sql="select * from [information] "&wh&" order by id desc"
'这里我们用到了 wh 变量，如果已确定按类别进行查询，实际的 SQL 语句如下行，否则值为空，不影响任何操作,这样可以避免因多次查询所带来的麻烦
'sql="select * from [information] where bigclassid="&request("bigclassid")&" and classid="&request("classid")&" order by id desc" 
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
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
		if rs("shuliang")<20 then
			response.Write("<script>alert('商品 "&rs("mingcheng")&" 数量已经少于 20！');</script>")	'判断每个商品的库存量是否有20，否则给出提示
		end if
%>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td><a href="pro.asp?id=<%=rs("id")%>"><%=rs("mingcheng")%></a></td>
          <td><%=rs("riqi")%></td>
          <td><%=rs("shichang")%></td>
          <td><%=rs("huiyuan")%></td>
          <td><a href="uppro.asp?id=<%=rs("id")%>">修改</a></td>
          <td><a href="lookpro.asp?id=<%=rs("id")%>&action=del&bigclassid=<%=request("bigclassid")%>&classid=<%=request("classid")%>">删除</a></td>
        </tr>
<%
	rs.movenext
	if rs.eof then exit for
	next
	end sub
%>
        <tr height="20" bgcolor="#FFFFFF" align="center">
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
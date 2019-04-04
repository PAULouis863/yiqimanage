
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd"> 

<% 
nowfilename=replace(replace(replace(now,":","")," ",""),"/","")
Response.Buffer = True 
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition", "inline; filename = "&nowfilename&".xls"
%> 
<html> 
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Untitled Document</title> 
</head> 
<body>
<%
if request("bigsort")<>"" and request("smallsort")<>"" then	'判断是否按指定类别进行查询
	wh="where bigclassid="&request("bigsort")&" and classid="&request("smallsort")&""
end if
sql="select * from [information] "&wh&" order by id desc"
'这里我们用到了 wh 变量，如果已确定按类别进行查询，实际的 SQL 语句如下行，否则值为空，不影响任何操作,这样可以避免因多次查询所带来的麻烦
'sql="select * from [information] where bigclassid="&request("bigclassid")&" and classid="&request("classid")&" order by id desc" 

		set rs=Server.CreateObject("ADODB.RECORDSET")

rs.open sql,conn,1,1
%>
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0">
<tr>
 
  <th>仪器名称</th>
  <th>上架时间</th>
  <th>出租价格</th>
  <th>购买价格</th>
 
 
 
 
</tr>
<%
i=0
do while rs.eof=false

%>
<tr class="TD2">

<td align="center"><%=rs("mingcheng")%></td>
<td align="center"><%=rs("riqi")%></td>
<td align="center"><%=rs("shichang")%></td>
<td align="center"><%=rs("huiyuan")%></td>





</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</table>
</body>
</html>

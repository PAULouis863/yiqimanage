
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
sql="select * from [user] order by id desc;"


		set rs=Server.CreateObject("ADODB.RECORDSET")

rs.open sql,conn,1,1
%>
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0">
<tr>
   <th>用户I D</th>
   <th>姓名</th>
  <th>电子邮件</th>
 <th>电话</th>
 <th>地址</th>
 
 
 
 
</tr>
<%
i=0
do while rs.eof=false

%>
<tr class="TD2">

            

  <td align="center"><%=rs("name")%></td>
            <td align="center"><%=rs("xingming")%></td>
            <td align="center"><%=rs("mail")%></td>
            <td align="center"><%=rs("tel")%></td>
            <td align="center"><%=rs("dizhi")%></td> 




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

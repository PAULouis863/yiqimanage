<!--#include file="../include/md5.asp" -->
<%
if session("admin")="" then
response.Write("<script>alert('非法操作！');window.location.href='login.asp';</script>")
else
sql="select * from [admin] where username='"&session("admin")&"';"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then
if rs("pass")<>session("admin_pass") then
session("admin")=""
response.Write("<script>alert('非法操作！');window.location.href='login.asp';</script>")
response.End()
end if
else
session("admin")=""
response.Write("<script>alert('非法操作！');window.location.href='login.asp';</script>")
response.End()
end if
end if
%>
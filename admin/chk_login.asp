<!--#include file="../include/conn.asp" -->
<!--#include file="../include/md5.asp" -->
<!--#include file="../include/include.asp" -->
<%
if request("logout")<>"" then
session("admin")=""
response.Write("<script>alert('注销成功！');window.location.href='../index.asp';</script>")
end if
if request("action")="login" then
sql="select * from [admin] where username='"&trim(request("username"))&"';"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then
if rs("pass")=md5(trim(request("pass"))) then
session("admin")=trim(request("username"))
session("admin_pass")=md5(trim(request("pass")))
response.Redirect("manage.asp")
else
session("admin")=""
session("admin_pass")=""
response.Write("<script>alert('用户名或密码错误！');window.location.href='login.asp';</script>")
end if
else
session("admin")=""
session("admin_pass")=""
response.Write("<script>alert('用户名或密码错误！');window.location.href='login.asp';</script>")
end if
end if
response.Write("<script>alert('用户名或密码错误！');window.location.href='login.asp';</script>")
%>
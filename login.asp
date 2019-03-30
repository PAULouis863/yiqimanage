<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="include/conn.asp" -->
<!--#include file="include/md5.asp" -->
<%
if request("login")="out" then	
	session("cishu")=""
	session("shijian")=""
	session("user")=""
	response.Redirect("index.asp")	
	response.End()
end if
sql="select * from [user] where name='"&trim(request("user"))&"';"	
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then	
	if rs("pass")=md5(trim(request("pass"))) then
		session("shijian")=rs("shijian2")
		session("cishu")=rs("cishu")
		session("user")=trim(request("user"))	
		rs("shijian2")=now()	
		rs("cishu")=rs("cishu")+1	
		rs.update
		response.Redirect("index.asp")	
	else	
		session("user")=""
		session("cishu")=""
		session("shijian")=""
		response.Write("<script>alert('用户名或密码错误！');window.location.href='index.asp';</script>")
		response.End() 
	end if
else	
	session("user")=""
	session("cishu")=""
	session("shijian")=""
	response.Write("<script>alert('用户名或密码错误！');window.location.href='index.asp';</script>")
	response.End() 
end if
rs.close
set rs=nothing
%>

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


		set rs=Server.CreateObject("ADODB.RECORDSET")

rs.open sql,conn,1,1
%>
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="0">
<tr>
 
  <th>订单编号</th>
  <th>收货方式</th>
  <th>付款方式</th>
  <th>下单用户</th>
  <th>下单时间</th>
  
 
 
 
</tr>
<%
i=0
do while rs.eof=false

%>
<tr class="TD2">
 <td align="center"><%=rs("didanhao")%></td>
            <td align="center"><%=rs("songhuo")%></td>
            <td align="center"><%=rs("zhifu")%></td>
            <td align="center"><%=rs("name")%></td>
            <td align="center"><%=rs("shijian")%></td>
            






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

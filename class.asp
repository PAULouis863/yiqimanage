<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.style2 {color: #000000}
-->
</style>
<!--#include file="include/conn.asp" -->
<!--#include file="include/include.asp" -->

<table width="96%" height="153"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
<%
if request("id")<>"" then
sql="select * from [information] where classid="&request("id")&" order by id desc"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof And rs.bof then
	Response.Write "<p align='center' class='contents'> 该分类下暂无！</p>"
else
	rs.pagesize=8
	SafeRequest(request("page"))
	page=clng(request("page"))
	if page<1 then page=1
	if page>rs.pagecount then page=rs.pagecount
	show rs,page
	sub show(rs,page)
	rs.absolutepage=page
	for i=1 to rs.pagesize
%>
              <td height="89"><table width="255"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="130" rowspan="7"><div align="center"><a href="lookpro.asp?id=<%=rs("id")%>" target="_blank"><img src="upfile/<%=rs("tupian")%>" width="110" height="130" border="0"></a></div></td>
                    <td width="20" height="16">&nbsp;</td>
                    <td width="113"><span class="style2">【<%=rs("mingcheng")%>】</span></td>
                  </tr>
                  <tr>
                    <td height="16">&nbsp;</td>
                    <td><span class="style2">【出租价格：<%=rs("shichang")%>】</span></td>
                  </tr>
                 
                  <tr>
                    <td height="16">&nbsp;</td>
                    <td><a href="lookpro.asp?id=<%=rs("id")%>" target="_blank">【查看信息</a>】</td>
                  </tr>
           
				  <tr>
                    <td height="16">&nbsp;</td>
                    <td>【<a href="#" onClick='javascript:parent.window.location.href="gouwu.asp?ProdId=<%=rs("id")%>&class=class";'>租借影片</a>】</td>
                  </tr>
                  <tr>
                    <td height="16">&nbsp;</td>
                    <td><font color="13589B">【浏览次数：<%=rs("cishu")%>】</font></td>
                  </tr>
              </table></td>
	<%if i mod 2=0 then%>
            </tr>
<tr><td height="10"></td>
</tr>
<%
end if
	rs.movenext
	if rs.eof then exit for
	next
	end sub
%>
      <tr>
	  <form action='' method='get' name='form'>
        <td height="30" colspan="2">
          <div align="center">
<%	
	if page<>1 then
		response.Write("&nbsp;&nbsp;<a href="&path&"?page=1&id="&request("id")&">第一页</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&(page-1)&"&id="&request("id")&">上一页</a>")
	end if 
	response.Write("&nbsp;&nbsp;当前 <font color='#FF0000'>"&page&"</font> 页")
	response.Write("&nbsp;&nbsp;条 <font color='#FF0000'>"&rs.recordcount&"</font>/<font color='#FF0000'>"&rs.pagecount&"</font> 页")
	if page<>rs.pagecount then
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&(page+1)&"&id="&request("id")&">下一页</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&rs.pagecount&"&id="&request("id")&">最末页</a>")
	end if
	response.Write("&nbsp;&nbsp;跳转到<input type='text' size='2' name='page'>页<input type='hidden' name='id' value='"&request("id")&"'><input type='submit' value='GO'>")
end if
rs.close
set rs=nothing
else
	response.Write("请选择类别查询！")
end if
%>
	      </div></td>
        </form></tr>
</table>
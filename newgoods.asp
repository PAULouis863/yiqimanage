<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.style1 {color: #f2ab5b}
.style2 {color: #000000}
-->
</style>
<!--#include file="include/conn.asp" -->
<!--#include file="include/include.asp" -->
<table width="792" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="3"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="792" height="165" background="images/index_r1_c1.jpg"><!--#include file="top.asp" --></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td width="197"><!--#include file="left.asp" --></td>
    <td width="590" valign="top"><table border="0" cellpadding="0" cellspacing="0" width="590">
      <tr>
        <td colspan="3"><img name="index_7_r1_c1" src="images/index_7_r1_c1.jpg" width="590" height="49" border="0" alt=""></td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r2_c1" src="images/index_7_r2_c1.jpg" width="590" height="9" border="0" alt=""></td>
      </tr>
      <tr>
        <td width="16" height="643" background="images/index_7_r3_c1.jpg">&nbsp;</td>
        <td width="565" valign="top"><table width="96%" height="153"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
<%
sql="select * from [information] order by id desc"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof And rs.bof then
	Response.Write "<p align='center' class='contents'> 对不起，暂无内容！</p>"
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
                    <td><span class="style2"><a href="lookpro.asp?id=<%=rs("id")%>" target="_blank">【查看信息</a>】</span></td>
                  </tr>
              
				  <tr>
                  <td height="16">&nbsp;</td>
                    <td><span class="style2">【<a href="gouwu.asp?ProdId=<%=rs("id")%>">租借影片</a>】</span></td>
                </tr>
                  <tr>
                    <td height="16">&nbsp;</td>
                    <td><span class="style2">【浏览次数：<%=rs("cishu")%>】</span></td>
                  </tr>
              </table></td>
	<%if i mod 2=0 then%>
            </tr>
<tr><td height="22"></td>
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
		    </div></td>
        </form></tr>
        </table></td>
        <td width="10" background="images/index_7_r3_c3.jpg">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r4_c1" src="images/index_7_r4_c1.jpg" width="590" height="7" border="0" alt=""></td>
      </tr>
    </table></td>
    <td width="8"><img name="index_r2_c3" src="images/index_r2_c3.jpg" width="5" height="753" border="0" alt=""></td>
  </tr>
  <tr>
    <td colspan="3"><!--#include file="foot.asp" --></td>
  </tr>
</table>
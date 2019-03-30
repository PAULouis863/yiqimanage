<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.style1 {color: #f2ab5b}
-->
</style>
<!--#include file="include/conn.asp" -->
<!--#include file="include/include.asp" -->
<!--#include file="chk_user.asp" -->
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
        <td colspan="3"><img name="index_7_r1_c1" src="images/gwc.jpg" width="590" height="49" border="0" alt=""></td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r2_c1" src="images/index_7_r2_c1.jpg" width="590" height="9" border="0" alt=""></td>
      </tr>
      <tr>
        <td width="16" height="643" background="images/index_7_r3_c1.jpg">&nbsp;</td>
        <td width="565" valign="top">
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<%
if request("ProductList")="ProductList" then	
	Session("ProductList")=""
	response.Write("<script>alert('您的购物车为空!');window.location.href='index.asp';</script>")
end if

ProductList = Session("ProductList")	
Products = Split(Request("Prodid"), ",")	
For I=0 To UBound(Products)	
   PutToShopBag Products(I), ProductList	
Next
Session("ProductList") = ProductList	

Sub PutToShopBag( Prodid, ProductList )	
   If Len(ProductList) = 0 Then	
      ProductList =Prodid	
   ElseIf InStr( ProductList, Prodid ) <= 0 Then	
      ProductList = ProductList&", "&Prodid &""	
   End If
End Sub

If Request("update") = "update" Then	
   ProductList = ""	
   Products = Split(Request("ProdId"), ", ")	
   For I=0 To UBound(Products)	
      PutToShopBag Products(I), ProductList	
   Next
   Session("ProductList") = ProductList	
End If

If Len(Session("ProductList")) = 0 Then
	response.Write("<script>alert('您的购物车为空!');window.location.href='index.asp';</script>")
	Response.end
end if
%>
<table width="96%" border="0" cellspacing="0" cellpadding="0" align="center">
<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
strsql="select * from information where ID in ("&Session("ProductList")&") order by ID"	
rs.open strsql,conn,1,1
%>
<tr> <td>
<form action="gouwu.asp" method="POST" name="check">
      <table border="0" cellspacing="1" cellpadding="4" align="center" width="100％" bgcolor="BDBDBC">
            <tr bgcolor="#FFFFFF" height="25" align="center"> 
            <td width="40">编 号</td>
            <td width="300">影 片 名 称</td>
            <td width="40">数量</td>
			 <td width="60">出租价格</td>
			
            <td width="60">成交价</td>
			<td width="70">小 计</td>
          </tr>
<%
Sum = 0	
Quatity = 1	
Do While Not rs.EOF
	Quatity = Request.Form( "Q_" & rs("ID"))	
	If Quatity <= 0 Then	
		Quatity = Session(rs("ID"))
		If Quatity <= 0 Then Quatity = 1	
	End If
	Session(rs("ID")) = Quatity	
	Sum = Sum + rs("shichang")*Quatity	
%> 
          <tr bgcolor="#FFFFFF" height="25"align="center"> 
            <td><input type="CheckBox" name="ProdId" value="<%=rs("ID")%>" Checked></td>
			 <input type="hidden" name="shuliang" value="<%response.Write Quatity	%>">
            <td align="left">&nbsp;<a href="lookpro.asp?ID=<%=rs("ID")%>" target="_blank"><%=rs("mingcheng")%></a></td>
            <td><input type="Text" name="<%response.Write("Q_" & rs("ID")) %>" value="<%response.Write Quatity	%>" size="2" class="form"></td>
			<td><%=rs("shichang")%></td>
			<td><%=rs("shichang")%></td>
			<td><%=rs("shichang")%></td>
			<td><%response.Write(rs("shichang")*Quatity)	%></td>
          </tr>
		  <input type="hidden" name="xiaoji" value="<%response.Write(rs("huiyuan")*Quatity)	%>">
<%
	rs.MoveNext
	Loop
rs.close
set rs=nothing
%> 
<tr bgcolor="#FFFFFF"> 
 <td height="25" colspan="8" align="center" valign="middle">             
                <input type="submit" name="order" value="更新影片"> &nbsp;&nbsp;&nbsp;
                <input type="reset" name="payment" value="去收银台" onClick="window.location.href='shouyin.asp';"> 
                &nbsp;&nbsp;&nbsp; 
				&nbsp;&nbsp;<a href="gouwu.asp?ProductList=ProductList">清空购物车</a>&nbsp;&nbsp;&nbsp;&nbsp;总计：<%=Sum%>
                <input type="hidden" name="update" value="update">
</td>
</tr>
      </table>
      </form>
 </td>
</tr>
      </table>
	  		</td>
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
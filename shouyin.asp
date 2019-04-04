<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.style1 {color: #f2ab5b}
-->
</style>
<!--#include file="include/conn.asp"-->
<!--#include file="chk_user.asp" -->
<!--#include file="include/include.asp" -->
<%
if request("ProductList")="ProductList" then	
	Session("ProductList")=""
	response.Write("<script>alert('您的购物车为空!');window.location.href='index.asp';</script>")
	response.End()
end if
%>
<SCRIPT>
<!--
function chk()
{
   if(document.receiveaddr.shoujianren.value=="") 
	{
		document.receiveaddr.shoujianren.focus();
		alert("请填写收货人姓名！");
		return false;
	}
	
   if(document.receiveaddr.dizhi.value=="") 
	{
		document.receiveaddr.dizhi.focus();
		alert("请填写收货人地址！");
		return false;
	}
	
   if(document.receiveaddr.youbian.value=="") 
	{
		document.receiveaddr.youbian.focus();
		alert("请填写邮编！");
		return false;
	}

   if(document.receiveaddr.tel.value=="") 
	{
		document.receiveaddr.tel.focus();
		alert("请填写联系电话！");
		return false;
	}

   if(document.receiveaddr.mail.value=="") 
	{
		document.receiveaddr.mail.focus();
		alert("请填写电子邮件！");
		return false;
	}

}
//-->
</script> 
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
        <td colspan="3"><img name="index_7_r1_c1" src="images/syt.jpg" width="590" height="49" border="0" alt=""></td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r2_c1" src="images/index_7_r2_c1.jpg" width="590" height="9" border="0" alt=""></td>
      </tr>
      <tr>
        <td width="16" height="643" background="images/index_7_r3_c1.jpg">&nbsp;</td>
        <td width="565" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table>
<%
if request("action")="add" then




sqls="select * from [user] where name='"&session("user")&"';"
	set rss=Server.CreateObject("ADODB.Recordset")
	rss.open sqls,conn,3,3
	rss("xf")=rss("xf")+session("xf")

	rss.update
	rss.close
	set rss=nothing















	sql="select * from [order]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("name")=session("user")	
	rs("information")=request("ProdId")	
	rs("shijian")=now()
	rs("shuliang")=request("shuliang")
	rs("mail")=request("mail")
	rs("dizhi")=trim(request("dizhi"))
	rs("youbian")=trim(request("youbian"))
	rs("zhuangtai")="0"	
	rs("tel")=trim(request("tel"))
	rs("shoujianren")=trim(request("shoujianren"))
	rs("zhifu")=trim(request("zhifu"))
	rs("songhuo")=trim(request("songhuo"))
	rs("didanhao")=GetOrderNo(Now())	
	rs("leaveword")=trim(request("leaveword"))
    rs("xf")=session("je")
    
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('定货成功！\n\n 牢记您的的订单号"&GetOrderNo(Now())&"');window.location.href='shouyin.asp?ProductList=ProductList';</script>")
end if
%>
<form name="receiveaddr" method="post" action="shouyin.asp" onSubmit="return chk();">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
strsql="select * from information where ID in ("&Session("ProductList")&") order by ID"
rs.open strsql,conn,1,1
%>
<tr> <td>
      <table border="0" cellspacing="1" cellpadding="4" align="center" width="100％" bgcolor="BDBDBC">
            <tr bgcolor="#FFFFFF" height="25" align="center"> 
            <td width="40">编 号</td>
            <td width="300">仪 器 名 称</td>
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
    session("xf")=Sum
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
             
           &nbsp;&nbsp;&nbsp;    &nbsp;&nbsp;&nbsp;    &nbsp;&nbsp;&nbsp;   &nbsp;&nbsp;&nbsp;    &nbsp;&nbsp;&nbsp;    &nbsp;&nbsp;&nbsp;   &nbsp;&nbsp;&nbsp;    &nbsp;&nbsp;&nbsp;    &nbsp;&nbsp;&nbsp;    &nbsp;&nbsp;&nbsp;    &nbsp;&nbsp;&nbsp;    &nbsp;&nbsp;&nbsp;  &nbsp;&nbsp;&nbsp;    &nbsp;&nbsp;&nbsp;    &nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;<%if session("je")<100 Then response.Write("铜牌会员九折 总计："&Sum*0.9) else if  session("je")>100 and session("je")<500 Then response.Write("银牌会员八折 总计："&Sum*0.8)   else if   session("je")>500 Then response.Write("金牌会员七折 总计："&Sum*0.7)    end if %>
                <input type="hidden" name="update" value="update">
</td>
</tr>
      </table>
 </td>
</tr>
  </table>
<%
sql="select * from [user] where name='"&session("user")&"';"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then
	xingming=rs("xingming")
	dizhi=rs("dizhi")
	tel=rs("tel")
	mail=rs("mail")
	youbian=rs("youbian")
end if
%>
<table width="96%" border="0" cellspacing="0" cellpadding="0" align="center" class="table-zuoyou"> 
<tr><td>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="BDBDBC">
          <tr bgcolor=#ffffff> 
		    <input type=hidden name=realname value=timesshop>
            <td width="150">收货人姓名：</td>
            <td width="600" height="28">  
              <input name="shoujianren" type="text" size="12" value=<%=xingming%>>
            </td>
          </tr>
          <tr bgcolor=#ffffff> 
            <td>详细地址：</td>
            <td height="28"><input name="dizhi" type="text" size="40" value=<%=dizhi%>></td>
          </tr>
          <tr bgcolor=#ffffff> 
            <td>邮　　编：</td>
            <td height="28"><input name="youbian" type="text"  size="10" value=<%=youbian%>></td>
          </tr>
          <tr bgcolor=#ffffff> 
            <td>电　　话：</td>
            <td height="28"><input name="tel" type="text" size="12" value=<%=tel%>></td>
          </tr>
          <tr bgcolor=#ffffff> 
            <td>电子邮件：</td>
            <td height="28"><input name="mail" type="text" value=<%=mail%>></td>
          </tr>
          <tr bgcolor=#ffffff> 
            <td>送货方式：</td>
            <td height="28">
          <select name=songhuo size=3 id=deliverymethord>
           <option value=本店交易> 本店交易</option> <option value=普通平邮>普通平邮</option><option value=特快专递>特快专递（EMS）</option><option value=送货上门 selected>送货上门</option></select></td>
          </tr>
          <tr bgcolor=#ffffff> 
            <td>支付方式：</td>
            <td height="20">
           <select name=zhifu size=4 id=paymethord><option value=工商银行汇款>工商银行汇款</option><option value=建设银行汇款 selected>建设银行汇款</option><option value=邮局汇款>邮局汇款</option><option value=交通银行汇款>交通银行汇款</option></select>
            </td>
          </tr>
          <tr bgcolor=#ffffff> 
            <td valign="top">简单留言：</td>
            <td height="28"><textarea name="leaveword" cols="40" rows="5" id="comments"></textarea></td>
          </tr>
          <tr bgcolor=#ffffff> 
		    <td></td>
            <td><input type="submit" name="Submit3" value="提交订单"></td>
          </tr>
		 <input type="hidden" name="action" value="add">
      </table> </td>
  </tr> </table></form>
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
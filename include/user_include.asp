<!--#include file="../include/md5.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
sub updateuser()	
	sql="select * from [user] where id="&request("id")&" and name='"&trim(session("user"))&"';"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs("mail")=trim(request("mail"))
	rs("youbian")=trim(request("youbian"))
	rs("xingming")=trim(request("xingming"))
	rs("shenfenzheng")=trim(request("shenfenzheng"))
	rs("tel")=trim(request("tel"))
	rs("qq")=trim(request("qq"))
	rs("tishi")=trim(request("tishi"))
	rs("huida")=trim(request("huida"))
	rs("dizhi")=trim(request("dizhi"))
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('修改成功！');window.location.href='usercenter.asp?action=1';</script>")
end sub

sub updatepass()	
	sql="select * from [user] where id="&request("id")&" and name='"&trim(session("user"))&"';"	
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs("pass")=md5(request("newpass1"))	
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('修改成功！');window.location.href='usercenter.asp?action=1';</script>")
end sub 

sub thispass()	
	sql="select * from [user] where name='"&request("name")&"';"	
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	if not rs.eof then	
		if rs("huida")=request("huida") and rs("tishi")=request("tishi") then	
			session("shijian")=rs("shijian2")	
			session("cishu")=rs("cishu")
			session("user")=trim(request("name"))
			rs("shijian2")=now()
			rs("cishu")=rs("cishu")+1
			rs.update
			response.Write("<script>alert('系统提示\n\n验证通过，进入自动登陆状态！');window.location.href='usercenter.asp?action=3';</script>")	'进入用户中心的密码修改模块中
			response.End()
		end if
	end if
	rs.close
	set rs=nothing
	response.Write("<script>alert('数据不正确，无法查询！');window.location.href='usercenter.asp?action=1';</script>")	'也许是用户名、提示、回答中任何一个输入不符（没有做具体判断是那种错误，因为不可以告诉用户过于详细的信息，避免无限的猜解）
end sub
%>

<%sub	lookuser()	%>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
<%
sql="select * from [user] where name='"&trim(session("user"))&"';"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then
%>
<form name="myform" action="lookuser.asp" method="post">  <tr> 
    <td align="center"><span class="style5">详细信息</span></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="21%" height="30">用&nbsp;户&nbsp;名：</td>
          <td width="79%"><div align="left"><%=rs("name")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">电子邮件：</td>
          <td><div align="left"><%=rs("mail")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">姓&nbsp;&nbsp;&nbsp;&nbsp;名：</td>
          <td><div align="left"><%=rs("xingming")%></div></td>
        </tr>

                <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">会员等级：</td>
          <td><div align="left"><%if rs("xf")<100 Then response.Write("铜牌会员") else if   rs("xf")>100 and rs("xf")<500 Then response.Write("银牌会员")   else if    rs("xf")>500 Then response.Write("金牌会员")    end if %></div></td>
        </tr>








        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">电&nbsp;&nbsp;&nbsp;&nbsp;话：</td>
          <td><div align="left"><%=rs("tel")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">身&nbsp;份&nbsp;证：</td>
          <td><div align="left"><%=rs("shenfenzheng")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">邮&nbsp;&nbsp;&nbsp;&nbsp;编：</td>
          <td><div align="left"><%=rs("youbian")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">地&nbsp;&nbsp;&nbsp;&nbsp;址：</td>
          <td><div align="left"><%=rs("dizhi")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">联系&nbsp;&nbsp;QQ：</td>
          <td><div align="left"><%=rs("qq")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">提&nbsp;&nbsp;&nbsp;&nbsp;示：</td>
          <td><div align="left"><%=rs("tishi")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">注册时间：</td>
          <td><div align="left"><%=rs("shijian1")%>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">上次登录：</td>
          <td><div align="left"><%=rs("shijian2")%>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">登录次数：</td>
          <td><div align="left"><%=rs("cishu")%>
          </div></td>
        </tr>
      </table>
      <br></td>
  </tr>
</form>
<%
end if
rs.close
set rs=nothing
%>
</table>
		  </td>
        </tr>
</table>
<%end sub%>

<%sub upuser()	%>
<script>
function chk()
{
	if (document.myform.mail.value=="")
	{
		document.myform.mail.focus();
		alert("请输入电子邮件！");
		return false;
	}
	
	if (document.myform.youbian.value=="")
	{
		document.myform.youbian.focus();
		alert("请输入邮编！");
		return false;
	}
	
	if (document.myform.xingming.value=="")
	{
		document.myform.xingming.focus();
		alert("请输入真实姓名！");
		return false;
	}
	
	if (document.myform.tel.value=="")
	{
		document.myform.tel.focus();
		alert("请输入联系电话！");
		return false;
	}
	
	if (document.myform.shenfenzheng.value=="")
	{
		document.myform.shenfenzheng.focus();
		alert("请输入身份证！");
		return false;
	}
	
	if (document.myform.dizhi.value=="")
	{
		document.myform.dizhi.focus();
		alert("请输入地址！");
		return false;
	}
	
	if (document.myform.qq.value=="")
	{
		document.myform.qq.focus();
		alert("请输入联系qq！");
		return false;
	}
	
	if (document.myform.tishi.value=="")
	{
		document.myform.tishi.focus();
		alert("请输入密码提示！");
		return false;
	}
	
	if (document.myform.huida.value=="")
	{
		document.myform.huida.focus();
		alert("请输入密码回答！");
		return false;
	}
	
}
</script>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
      <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
<%
sql="select * from [user] where name='"&trim(session("user"))&"';"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
'if not rs.eof then
%>
        <form name="myform" action="usercenter.asp" method="post"  onSubmit="return chk();">
          <tr>
            <td align="center"><span class="style5">修改信息</span></td>
          </tr>
          <tr>
            <td valign="top" bgcolor="#FFFFFF"><br>
                <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td width="21%">用&nbsp;户&nbsp;名：</td>
                    <td width="79%"><div align="left"><%=rs("name")%></div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>电子邮件：</td>
                    <td><div align="left">
                        <input name="mail" type="text" value="<%=rs("mail")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>邮&nbsp;&nbsp;&nbsp;&nbsp;编：</td>
                    <td><div align="left">
                        <input name="youbian" type="text" value="<%=rs("youbian")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>姓&nbsp;&nbsp;&nbsp;&nbsp;名：</td>
                    <td><div align="left">
                        <input name="xingming" type="text" value="<%=rs("xingming")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>电&nbsp;&nbsp;&nbsp;&nbsp;话：</td>
                    <td><div align="left">
                        <input name="tel" type="text" value="<%=rs("tel")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>身&nbsp;份&nbsp;证：</td>
                    <td><div align="left">
                        <input name="shenfenzheng" type="text" value="<%=rs("shenfenzheng")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>地&nbsp;&nbsp;&nbsp;&nbsp;址：</td>
                    <td><div align="left">
                        <input name="dizhi" type="text" value="<%=rs("dizhi")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>联系&nbsp;&nbsp;QQ：</td>
                    <td><div align="left">
                        <input name="qq" type="text" value="<%=rs("qq")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>提&nbsp;&nbsp;&nbsp;&nbsp;示：</td>
                    <td><div align="left">
                        <input name="tishi" type="text" value="<%=rs("tishi")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>回&nbsp;&nbsp;&nbsp;&nbsp;答：</td>
                    <td><div align="left">
                        <input name="huida" type="password" value="<%=rs("huida")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>注册时间：</td>
                    <td><div align="left"><%=rs("shijian1")%> </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>上次登陆：</td>
                    <td><div align="left"><%=rs("shijian2")%> </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>登陆次数：</td>
                    <td><div align="left"><%=rs("cishu")%> </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td colspan="2">
                      <input name="action" type="hidden" value="updateuser">
                      <input name="id" type="hidden" value="<%=rs("id")%>">
                      <input type="reset" name="reset" value="关闭" onClick='javascript:window.close();'>
&nbsp;&nbsp;
                    <input name="submit" type="submit" value="修改"></td>
                  </tr>
                </table>
                <br></td>
          </tr>
        </form>
<%

rs.close
set rs=nothing
%>
      </table>
    </td>
  </tr>
</table>
<%end sub%>

<%sub uppass()	%>
<script>
function chk()
{
	if (document.myform.newpass1.value=="")
	{
		document.myform.newpass1.focus();
		alert("请输新密码！");
		return false;
	}
	
	if (document.myform.newpass2.value=="")
	{
		document.myform.newpass2.focus();
		alert("请确认新密码！");
		return false;
	}
	
	if (document.myform.newpass1.value!=document.myform.newpass2.value)
	{
		document.myform.newpass2.focus();
		alert("两次输入的密码不一样！");
		return false;
	}
}
</script>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
      <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
<%
sql="select * from [user] where name='"&trim(session("user"))&"';"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then
%>
        <form name="myform" action="usercenter.asp" method="post" onSubmit="return chk();">
          <tr>
            <td align="center"><span class="style5">修改密码</span></td>
          </tr>
          <tr>
            <td valign="top" bgcolor="#FFFFFF"><br>
                <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td width="21%">用&nbsp;户&nbsp;名：</td>
                    <td width="79%"><div align="left"><%=trim(session("user"))%></div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>原&nbsp;密&nbsp;码：</td>
                    <td><div align="left">
                        <input name="pass" type="password" value="<%=rs("pass")%>" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>新 密 码：</td>
                    <td><div align="left">
                        <input name="newpass1" type="password" value="" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>确认密码：</td>
                    <td><div align="left">
                        <input name="newpass2" type="password" value="" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td colspan="2">
                      <input name="action" type="hidden" value="updatepass">
                      <input name="id" type="hidden" value="<%=rs("id")%>">
                      <input name="up" type="hidden" value="<%=request("up")%>">
&nbsp;&nbsp;
                    <input name="submit" type="submit" value="修改"></td>
                  </tr>
                </table>
                <br></td>
          </tr>
        </form>
<%
end if
rs.close
set rs=nothing
%>
      </table>
    </td>
  </tr>
</table>
<%end sub%>

<%sub lookpass()	%>
<script>
function chk()
{
	if (document.myform.name.value=="")
	{
		document.myform.name.focus();
		alert("请输入用户名！");
		return false;
	}
	
	if (document.myform.tishi.value=="")
	{
		document.myform.tishi.focus();
		alert("请输入密码提示！");
		return false;
	}
	
	if (document.myform.huida.value=="")
	{
		document.myform.huida.focus();
		alert("请输入密码回答！");
		return false;
	}
}
</script>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
      <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
        <form name="myform" action="usercenter.asp" method="post" onSubmit="return chk();"<%'验证是否已经填写所有数据%>>
          <tr>
            <td align="center"><span class="style5">取回密码</span></td>
          </tr>
          <tr>
            <td valign="top" bgcolor="#FFFFFF"><br>
                <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td width="21%">用&nbsp;户&nbsp;名：</td>
                    <td width="79%"><div align="left"><input type="text" name="name" value=""></div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td width="21%">密码提示：</td>
                    <td width="79%"><div align="left"><input name="tishi" type="text" value="" size="30"></div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>密码回答：</td>
                    <td><div align="left">
                        <input name="huida" type="password" value="" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td colspan="2">
                      <input name="action" type="hidden" value="thispass"<%'隐藏提交%>>
&nbsp;&nbsp;
                    <input name="submit" type="submit" value="取回"></td>
                  </tr>
                </table>
                <br></td>
          </tr>
        </form>
    </table></td>
  </tr>
</table>
<%end sub%>

<%sub order()	%>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
<%
sql="select * from [order] where name='"&trim(session("user"))&"';"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then
%>
  <tr>
    <td>
      <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
          <tr>
            <td align="center"><span class="style5">用户订单</span></td>
          </tr>
          <tr>
            <td valign="top" bgcolor="#FFFFFF"><br>
                <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td width="175"><div align="center">订单编号</div></td>
                    <td width="201"><div align="center">付款方式</div></td>
                    <td width="215"><div align="center">送货方式</div></td>
                    <td width="216"><div align="center">下单时间</div></td>
                  </tr>
				  <%do while not rs.eof%>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td><a href="chaxun.asp?order=<%=rs("didanhao")%>" target="_blank"><%=rs("didanhao")%></a></td>
                    <td><%=rs("zhifu")%></td>
                    <td><%=rs("songhuo")%></td>
                    <td><%=rs("shijian")%></td>
                  </tr>
				  <%
				  rs.movenext
				  loop
				  %>
                </table>
                <br></td>
          </tr>
        </form>
    </table></td>
  </tr>
<%
end if
rs.close
set rs=nothing
%>
</table>  
<%end sub%>
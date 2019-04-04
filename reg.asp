<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="include/conn.asp" -->
<!--#include file="include/include.asp" -->
<!--#include file="include/md5.asp" -->
<%
if request("action")="add" then	
	sql="select * from [user] where name='"&trim(request("user"))&"';"	
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	if not rs.eof then	
		response.Write("<script>alert('该用户名已经被注册！');history.back();</script>")	
		response.End()	
	end if
	rs.close
	set rs=nothing

	sql="select * from [user]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3	
	rs.addnew	
	rs("name")=trim(request("user"))
	rs("pass")=md5(trim(request("pass")))	
	rs("mail")=trim(request("mail"))
	rs("youbian")=trim(request("youbian"))
	rs("xingming")=trim(request("xingming"))
	rs("shenfenzheng")=trim(request("shenfenzheng"))
	rs("tel")=trim(request("tel"))
	rs("qq")=trim(request("qq"))
	rs("tishi")=trim(request("tishi"))
	rs("huida")=trim(request("huida"))
	rs("dizhi")=trim(request("dizhi"))
	rs("shijian1")=now()	
	rs("cishu")="0"	
    rs("dengji")="铜牌会员"	
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('注册成功！');window.location.href='index.asp';</script>")
end if
%>
<script>
function chk()
{
	if (document.myform.user.value=="")
	{
		document.myform.user.focus();
		alert("请输入用户名！");
		return false;
	}
	
	if (document.myform.pass.value=="")
	{
		document.myform.pass.focus();
		alert("请输入密码！");
		return false;
	}
	
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
        <td colspan="3"><img name="index_7_r1_c1" src="images/yhzc.jpg" width="590" height="49" border="0" alt=""></td>
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
		
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
      <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
        <form name="myform" action="reg.asp" method="post"  onSubmit="return chk();">
          <tr>
            <td align="center"><span class="style5">用户注册</span></td>
          </tr>
          <tr>
            <td valign="top" bgcolor="#FFFFFF"><br>
                <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="BDBDBC">
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td width="21%">用&nbsp;户&nbsp;名：</td>
                    <td width="79%"><div align="left">
                      <input name="user" type="text" id="user" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>密&nbsp;&nbsp;&nbsp;&nbsp;码：</td>
                    <td><div align="left">
                        <input name="pass" type="password" id="pass" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>电子邮件：</td>
                    <td><div align="left">
                        <input name="mail" type="text" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>邮&nbsp;&nbsp;&nbsp;&nbsp;编：</td>
                    <td><div align="left">
                        <input name="youbian" type="text" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>姓&nbsp;&nbsp;&nbsp;&nbsp;名：</td>
                    <td><div align="left">
                        <input name="xingming" type="text" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>电&nbsp;&nbsp;&nbsp;&nbsp;话：</td>
                    <td><div align="left">
                        <input name="tel" type="text" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>身&nbsp;份&nbsp;证：</td>
                    <td><div align="left">
                        <input name="shenfenzheng" type="text" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>地&nbsp;&nbsp;&nbsp;&nbsp;址：</td>
                    <td><div align="left">
                        <input name="dizhi" type="text" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>联系&nbsp;&nbsp;QQ：</td>
                    <td><div align="left">
                        <input name="qq" type="text" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>密码提示：</td>
                    <td><div align="left">
                        <input name="tishi" type="text" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td>问题回答：</td>
                    <td><div align="left">
                        <input name="huida" type="password" size="30">
                    </div></td>
                  </tr>
                  <tr height="20" bgcolor="#FFFFFF" align="center">
                    <td colspan="2">
                      <input name="action" type="hidden" value="add">
                      <input type="reset" name="reset" value="重写" >
&nbsp;&nbsp;
                    <input name="submit" type="submit" value="注册"></td>
                  </tr>
                </table>
                <br></td>
          </tr>
        </form>
      </table>
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
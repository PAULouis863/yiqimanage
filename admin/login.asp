<!--#include file="../include/conn.asp" -->
<script language = "JavaScript">
function login()
{
	if(document.myform.username.value=="") 
	{
		document.myform.username.focus();
		alert("请输入用户名！");
		return false;
	}
	if(document.myform.pass.value=="") 
	{
		document.myform.pass.focus();
		alert("请输入密码！");
		return false;
	}
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.style2 {color: #FFFFFF}
-->
</style>
<table width="760" height="570"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td background="../images/houtai.jpg"><div align="center">

  <table width="100%"  border="0" cellspacing="0" cellpadding="0"><form name="myform" method="post" action="chk_login.asp" onSubmit="return login();">
    <tr>
      <td width="42%" height="26"><div align="right" class="style2"><strong>用户名：</strong></div></td>
      <td width="27%"><input name="username" type="text"></td>
      <td width="31%">&nbsp;</td>
    </tr>
    <tr>
      <td height="26"><div align="right" class="style2"><strong>密&nbsp;&nbsp;码：</strong></div></td>
      <td><input name="pass" type="password"></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td height="46">&nbsp;</td>
      <td><input name="action" type="hidden" value="login">
        <input name="submit" type="submit" value="登录">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input name="" type="reset" value="取消"></td>
      <td>&nbsp;</td>
    </tr>
  </form></table>
  <br>
  <br>
	</div></td>
  </tr>
</table>
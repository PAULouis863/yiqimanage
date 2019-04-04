<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
body  { background:#799AE1; margin:0px; font:9pt ??ì?; }
table  { border:0px; }
td  { font:normal 12px ??ì?; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px ??ì?; color:#000000; text-decoration:none; }
a:hover  { color:#428EFF;text-decoration:underline; }
.sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden; background:#D6DFF7; }
.menu_title  { }
.menu_title span  { position:relative; top:2px; left:8px; color:#215DC6; font-weight:bold; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:2px; left:8px; color:#428EFF; font-weight:bold; }
.style1 {color: #FF0000}
</style>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		<td valign="bottom">
		<img height="38" src="images/title.gif" width="158" border="0"></td>
	</tr>
	<tr>
		
    <td class="menu_title" background="images/title_bg_quit.gif" height="25"> 
      <span><a href="#" onClick='javascript:parent.window.location.href="../index.asp"; '><b>系统首页</b></a>| <A HREF="chk_login.asp?logout=yes" TARGET="_top"><B>注销登录</B></A></span></td>
	</tr>
	<tr>
		<td align="center">
		<font face="Webdings" color="#FFFFFF" style=cursor:hand>5</font> </td>
	</tr>
</table>


<table cellspacing="0" cellpadding="0" width="158" align="center">
<tr>		
    <td id="imgmenu1" class="menu_title" onmouseover="this.className='menu_title2';" onclick="showsubmenu(1)" onmouseout="this.className='menu_title';" style="cursor:hand" background=images/menudown.gif height="25"> 
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span>仪器管理</span></td>
  </tr>
	<tr>
		<td id="submenu1" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellpadding=0 cellspacing=0 align=center width=135> 
<tr><td height=20><div align="center"><a href=addpro.asp target=right>添加仪器信息</a></div></td></tr> 
<tr> <td height=20><div align="center"><a href=lookpro.asp target=right>仪器信息管理</a></div></td></tr> 
<tr><td height=20><div align="center"><a href=order.asp target=right>仪器订单管理</a></div></td></tr> 
<tr><td height=20><div align="center"><a href=comment.asp target=right>仪器评论管理</a></div></td></tr> 
</table>
		</div>
		<br>
		</td>
	</tr>
	<tr>
    <td class="menu_title" id="imgmenu2" onmouseover="this.className='menu_title2';" onclick="showsubmenu(2)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25"> 
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span>用户管理</span> </td>
	</tr>
	<tr>
		<td id="submenu2" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<div align="center">
				 <table cellpadding=0 cellspacing=0 align=center width=135>
          <tr> 
            <td height=20><div align="center"><a href=user.asp target=right>用户信息管理</a></div></td>
          </tr>
          <tr> 
            <td height=20><div align="center"><a href=master.asp target=right>后台管理员管理</a></div></td>
          </tr>
        </table>
			</div>
		</div>
		<br>
		</td>
	</tr>

	<tr>	
    <td class="menu_title" id="imgmenu3" onmouseover="this.className='menu_title2';" onclick="showsubmenu(3)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25"> 
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span>分类管理</span> </td>
	</tr>
	<tr>
		<td id="submenu3" style="DISPLAY: none" align="center">
		<div class="sec_menu" style="WIDTH: 158px">
<table cellpadding=0 cellspacing=0 align=center width=135> 
<tr> <td height=20><div align="center"><a href=bigclass.asp target=right>仪器大类管理</a></div></td></tr> 
<tr> <td height=20><div align="center"><a href=class.asp target=right>仪器小类管理</a></div></td></tr> 
</table>
		  </div><br>
		</td>
	</tr>
<tr>	
    <td class="menu_title" id="imgmenu4" onmouseover="this.className='menu_title2';" onclick="showsubmenu(4)" onmouseout="this.className='menu_title';" style="cursor:hand" background="images/menudown.gif" height="25"> 
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span>信息管理</span> </td>
  </tr>
	<tr>
	  <td id="submenu4" style="DISPLAY: none" align="center">
		<div class="sec_menu" style="WIDTH: 158px">
<table cellpadding=0 cellspacing=0 align=center width=136> 
<tr><td height=20><div align="center"><a href=notify.asp target=right>站内公告设置</a></div></td></tr> 
<tr><td height=20><div align="center"><a href=news.asp target=right>站内新闻管理</a></div></td></tr> 
<tr><td height=20><div align="center"><a href=addnews.asp target=right>添加站内新闻</a></div></td></tr> 
<tr><td height=20><div align="center"><a href=dismess.asp target=right>意见反馈管理</a></div></td></tr> 
<tr><td height=20><div align="center"><a href=book.asp target=right>留言板块管理</a></div></td></tr>
</table></div>
		</td>
	</tr>
</table>
<script language=javascript>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
imgmenu = eval("imgmenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
imgmenu.background="images/menuup.gif";
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
imgmenu.background="images/menudown.gif";
}
} 
</script>
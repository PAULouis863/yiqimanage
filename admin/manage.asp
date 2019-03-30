<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理</title>
<script>
if(self!=top){top.location=self.location;}
function switchSysBar()
{
	if (switchPoint.innerText==3)
	{
		switchPoint.innerText=4
		document.all("frmTitle").style.display="none"
	}
	else
	{
		switchPoint.innerText=3
		document.all("frmTitle").style.display=""
	}
}
if(window.screen.width<'1024'){switchSysBar()}
</script>
<style type="text/css">
	.navPoint {COLOR: white; CURSOR: hand; FONT-FAMILY: Webdings; FONT-SIZE: 9pt}
</style>
</head>
<body style="MARGIN: 0px" scroll=no onResize=javascript:parent.carnoc.location.reload()>
<table border="0" cellPadding="0" cellSpacing="0" height="100%" width="100%">
  <tr>
    <td align="middle" noWrap vAlign="center" id="frmTitle">
    <iframe frameBorder="0" id="left" name="left" scrolling=auto src="menu.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 170px; Z-INDEX: 2"></iframe>
    </td>
    <td bgcolor="A4B6D7" style="WIDTH: 9pt">
    <table border="0" cellPadding="0" cellSpacing="0" height="100%">
      <tr>
        <td style="HEIGHT: 100%" onclick="switchSysBar()">
        <font style="FONT-SIZE: 9pt; CURSOR: default; COLOR: #ffffff">
        <span class="navPoint" id="switchPoint" title="">3</span><br>
        </font></td>
      </tr>
    </table>
    </td>
    <td bgcolor="#FFFFFF" style="WIDTH: 100%">
<iframe frameBorder="0" id="right" name="right" scrolling="auto" src="admin.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1"></iframe>
    </td>
  </tr>
</table>
</body>
</html>  
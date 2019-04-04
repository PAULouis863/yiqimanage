<%
AccessDatabaseName="/database/shop.mdb"				  '数据库路径
AccessPassword=""
Dim ConnStr
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password="&AccessPassword&";Data Source=" & Server.MapPath(AccessDatabaseName)
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open ConnStr
%>

<head>
<title>仪器中心管理系统</title>
<script>top.document.title="仪器中心管理系统";</script>
<style>
BODY {
	font-family: "宋体";
	font-size: 9pt;
	font-style: normal;
	line-height: 160%;
	color: #000000;
	background-color: #FFFFFF;
}
TABLE {
	font-family: "宋体";
	font-size: 9pt;
	font-style: normal;
}
A:link {
	FONT-SIZE: 12px; COLOR: #000000; TEXT-DECORATION: none
}
A:visited {
	FONT-SIZE: 12px; COLOR: #000000; TEXT-DECORATION: none
}
A:active {
	FONT-SIZE: 12px; COLOR: #215DC6; TEXT-DECORATION: none
}
A:hover {
	FONT-SIZE: 12px; COLOR: #215DC6; TEXT-DECORATION: none;position: relative; right: 0px; top: 1px
}
.style1 {color: #f2ab5b}
</style>
</head>


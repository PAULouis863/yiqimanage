<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨����</title>
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->

<script language = "JavaScript">

var onecount;
onecount=0;
subcat = new Array();	
<%
i=0
sql="select * from [class] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof			''��ѯ����ѭ��������з���
	i=i+1	''���������±꣬���ʱASP���� i ��ֵΪ 1���������±��ʼֵΪ 0����������� i-1 ����Ϊ�˷��������±�Ĺ��� 
			''���ݿ����ж��ٷ��ϵ����ݣ���������������ж��٣���������ֵ���Ǳ������������ 1����Ϊ�����±��� 0 ��ͷ�� 
%>
	subcat[<%=i-1%>] = new Array("<%=rs("mingcheng")%>","<%=rs("bigclassid")%>","<%=rs("id")%>");
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
onecount=<%response.Write i ''����������������ȻASP���� i ��ѭ�����⣬�� i ��ѭ�������Ѿ���������ֵ%>;

function changelocation(locationid)
{
	document.myform.classid.length = 0;
	var locationid=locationid;	
	var i;
	for (i=0;i < onecount; i++)		
	{
		if (subcat[i][1] == locationid)	
		{ 
			 document.myform.classid.options[document.myform.classid.length] = new Option(subcat[i][0], subcat[i][2]);
			 
		}        
	}
}    

</script>

<script language = "JavaScript">
function addpro()
{
	if(document.myform.jianjie.value=="") 
	{
		document.myform.jianjie.focus();
		alert("������������飡");
		return false;
	}
	if(document.myform.riqi.value=="") 
	{
		document.myform.riqi.focus();
		alert("������������ڣ�");
		return false;
	}
	if(document.myform.mingcheng.value=="") 
	{
		document.myform.mingcheng.focus();
		alert("�������������ƣ�");
		return false;
	}
	if(document.myform.shichang.value=="") 
	{
		document.myform.shichang.focus();
		alert("�������г��۸�");
		return false;
	}
	if(document.myform.huiyuan.value=="") 
	{
		document.myform.huiyuan.focus();
		alert("�������Ա�۸�");
		return false;
	}
	if(document.myform.xinghao.value=="") 
	{
		document.myform.xinghao.focus();
		alert("�����������ͺţ�");
		return false;
	}
	if(document.myform.file.value=="") 
	{
		document.myform.file.focus();
		alert("���ϴ�ͼƬ��");
		return false;
	}
	if(document.myform.shuoming.value=="") 
	{
		document.myform.shuoming.focus();
		alert("����������˵����");
		return false;
	}
	if(document.myform.beizhu.value=="") 
	{
		document.myform.beizhu.focus();
		alert("������������ע��");
		return false;
	}
	if(document.myform.shuliang.value=="") 
	{
		document.myform.shuliang.focus();
		alert("����������������");
		return false;
	}
}
</script>
</head>
<body>
<%
if request("action")="add" then	
	
	SafeRequest(trim(request("shuliang")))	
	SafeRequest(trim(request("shichang")))
	SafeRequest(trim(request("huiyuan")))
	sql="select * from [information]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("jianjie")=trim(request("jianjie"))
	rs("riqi")=request("riqi")
	rs("mingcheng")=request("mingcheng")
	rs("shichang")=request("shichang")
	rs("huiyuan")=request("huiyuan")
	rs("xinghao")=request("xinghao")
	rs("dengji")=request("dengji")
	rs("tupian")=request("file")
	rs("shuoming")=request("shuoming")
	rs("beizhu")=request("beizhu")
	rs("bigclassid")=request("bigclassid")
	rs("classid")=request("classid")
	rs("shuliang")=request("shuliang")
	rs("cishu")="1"
	rs.update
	rs.close
	set rs=nothing
	session("tupian")=""	
	response.Write("<script>alert('��ӳɹ���');window.location.href='addpro.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
<form name="myform" action="addpro.asp" method="post" onSubmit="return addpro();">  <tr> 
    <td align="center"><font color="#FFFFFF">����µ�����</font></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="25%">������飺</td>
          <td width="75%"><div align="left">
            <input name="jianjie" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>�ϼ����ڣ�</td>
          <td><div align="left">
            <input name="riqi" type="text" value="<%=now()%>" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>�������ƣ�</td>
          <td><div align="left">
            <input name="mingcheng" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>����۸�</td>
          <td><div align="left">
            <input name="shichang" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>����۸�</td>
          <td><div align="left">
            <input name="huiyuan" type="text" size="30">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>������ţ�</td>
          <td><div align="left">
            <input name="xinghao" type="text" size="30">
          </div></td>
        </tr>
		<tr bgcolor="#FFFFFF">
			<td rowspan="2" align="center">����ͼƬ��</td>
			<td><div align="left"><input id="file" type="text" name="file"/>
				<script>
				function backfn(fname){
					document.getElementById("file").value=fname;
				}
				</script></div>
			</td>
		</tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          
          <td><div align="left">
			<iframe width="300" height="24" scrolling="no" MARGINHEIGHT="0" MARGINWIDTH="0" name="upload" src="upload.asp"></iframe>
          </div></td>
        </tr>
		
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>�������ܣ�</td>
          <td><div align="left">
            <textarea cols="60" rows="8" name="shuoming"></textarea>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>������ע��</td>
          <td><div align="left">
            <textarea cols="60" rows="8" name="beizhu"></textarea>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td colspan="2">�����ȼ���
            <select name="dengji">
                <option value="2">��ͨ</option>
                <option value="1">��Ʒ</option>
            </select>
<%if request("bigclassid")="" and request("classid")="" then
%>
�������ࣺ

<select name="bigclassid" size="1" id="bigclassid" onChange="changelocation(document.myform.bigclassid.options[document.myform.bigclassid.selectedIndex].value)">

<%
sql="select * from [bigclass] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof	
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
����С�ࣺ
<select name="classid">
<%
sql="select * from [class] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof	
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
<%else	%>
�������ࣺ
<select name="bigclassid" size="1" id="bigclassid">
<%
sql="select * from [bigclass] where id="&request("bigclassid")&" order by paixu;"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof	
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
����С�ࣺ
<select name="classid">
<%
sql="select * from [class] where id="&request("classid")&" order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3	
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.close
set rs=nothing
%>
</select>
<%end if%>
&nbsp;&nbsp;����������<input name="shuliang" type="text" size="6">
<input name="action" type="hidden" value="add">
&nbsp;&nbsp;<input type="reset" name="reset" value="��д">
&nbsp;&nbsp;<input name="submit" type="submit" value="���"></td>
          </tr>
      </table>
      <br></td>
  </tr>
</form>  
</table>
</body>
</html>
<%
formName=request("formName")
EditName=request("EditName")
%>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}  
-->
</style>
<table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
<form name="form1" method="post" action="upload.asp" enctype="multipart/form-data">
    <tr align="center" valign="middle"> 
      <td align="left" id="upid" height="20" width="400" bgcolor="#FFFFFF"> 
        <input name="file1" type="file" class="tx1" style="width:200" value="" size="30">
        <input type="submit" name="Submit" value="ÉÏ´«">
        <input type="hidden" name="EditName" value="<%=EditName%>">
        <input type="hidden" name="FormName" value="<%=formName%>">
        <input type="hidden" name="act" value="uploadfile"></td>
    </tr>
</form>
  </table>
<%
if request("class")<>"class" then
	if session("user")="" then
		response.Write("<script>alert('ÇëÏÈµÇÂ¼£¡');window.location.href='index.asp';</script>")
	end if
else
	if session("user")="" then
		response.Write("<script>alert('ÇëÏÈµÇÂ¼£¡');parent.window.location.href='index.asp';</script>")	
	end if
end if
%>
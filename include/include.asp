<%

Function HTMLEncode(Char)
	Char=Replace(Char,chr(13)," ")
	Char=Replace(Char,chr(10) & chr(10),"<p></p>")
	Char=Replace(Char,chr(10),"<br>")
	HTMLEncode=Char
End Function

Function SafeRequest(Para)
	If Not IsNumeric(Para) Then
		Response.Write ("<script>alert('参数 " & Para & " 必须为数字型，请正确操作！');history.back();</script>")
		Response.End
	End If
End Function

Function GetOrderNo(dDate)
    GetOrderNo = RIGHT("0000"+Trim(Year(dDate)),4)+RIGHT("00"+Trim(Month(dDate)),2)+RIGHT("00"+Trim(Day(dDate)),2)+RIGHT("00" + Trim(Hour(dDate)),2)+RIGHT("00"+Trim(Minute(dDate)),2)+RIGHT("00"+Trim(Second(dDate)),2)
End Function

function filter_Str(InString)
	NewStr=Replace(InString,"'","''")
	NewStr=Replace(NewStr,"<","&lt")
	NewStr=Replace(NewStr,">","&gt")
	NewStr=Replace(NewStr,"chr(60)","&lt;")
	NewStr=Replace(NewStr,"chr(37)","&gt;")
	NewStr=Replace(NewStr,"""","&quot")
	NewStr=Replace(NewStr,";",";;")
	NewStr=Replace(NewStr,"--","-")
	NewStr=Replace(NewStr,"/*"," ")
	NewStr=Replace(NewStr,"%"," ")
	filter_Str=NewStr
end function
%> 
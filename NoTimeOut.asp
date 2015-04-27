<%
If session("user_name")<>"" Then
	Response.Write(now())
Else
	Response.Write("Out")
End If
%>
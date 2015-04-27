<!--#include file="Comm.asp"-->
<!--#include file="API.asp" -->
<%
sql="select * from [Music] whrer [IsPlay]=True order by [id] desc" 
set rs=server.createobject("adodb.recordset") 
rs.open sql,conn,1,1 

response.contenttype="text/xml" 
response.write "<?xml version=""1.0"" encoding=""utf-8"" ?>" & vbcrlf 
response.write "<xml>" & vbcrlf 
while not rs.eof 
	response.write "<track>" & vbcrlf 
	response.write "<path>" & rs("Url") & "</path>" & vbcrlf 
	response.write "<title>" & rs("Name") & "</title>" & vbcrlf 
	response.write "</track>" & vbcrlf 
	rs.movenext 
wend
response.write "</xml>" & vbcrlf 
rs.close 

set rs=nothing 
%> 

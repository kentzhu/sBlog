<!--#include file="Comm.asp"-->
<!--#include file="API.asp" -->
<% 
strURL = "http://" & request.servervariables("server_name") & _ 
left(request.servervariables("script_name"),len(request.servervariables("SCRIPT_NAME"))-len("/rss.asp"))
sql="select top " & GetInf("Num_DefaultArt") & " * from [Article] order by [time] desc" 
set rs=server.createobject("adodb.recordset") 
rs.open sql,conn,1,1 

response.contenttype="text/xml" 
response.write "<?xml version=""1.0"" encoding=""gb2312"" ?>" & vbcrlf 
response.write "<rss version=""2.0"">" & vbcrlf 
response.write "<channel>" & vbcrlf 
response.write "<title>" & GetInf("Page_Title_1") & "</title>" & vbcrlf 
response.write "<link>" & strURL & "</link>" & vbcrlf 
response.write "<language>zh-cn</language>" & vbcrlf 
response.write "<copyright>An RSS feed for " & GetInf("Blogger_Name") & "</copyright>" & vbcrlf 

while not rs.eof 
	response.write "<item>" & vbcrlf 
	response.write "<title><![CDATA[" & rs("title") & "]]></title>" & vbcrlf 
	response.write "<link>"&strURL&"/ArticleView.asp?id="&rs("id")&"</link>" & vbcrlf 
	response.write "<description><![CDATA[" & rs("content") & "<br /><br />]]></description>" & vbcrlf 
	response.write "<pubDate>" & return_RFC822_Date(rs("time"),"GMT") & "</pubDate>" & vbcrlf 
	response.write "</item>" & vbcrlf 
	rs.movenext 
wend 

response.write "</channel>" & vbcrlf 
response.write "</rss>" & vbcrlf 
rs.close 

set rs=nothing 



%> 

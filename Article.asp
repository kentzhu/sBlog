<!--#include file="Comm.asp" -->
<!--#include file="API.asp" -->
<%
r_ClassID=Request.QueryString("ClassID")
AnitInjection(r_ClassID)
acn=GetArticleClassName(r_ClassID)
PAGETITLE="文章列表 " & acn
Page_Description=PAGETITLE
Page_Keywords=acn

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"> 
<html xmlns="http://www.w3.org/1999/xhtml"> 
<!--#include file="C_Head.asp" -->
<body> 
<!--#include file="C_Header.asp" -->
<!-- start page --> 
<div id="page"> 
	<!-- start content --> 
	<div id="content">
		<!-- 输出文章列表 --> 
		<%Print_ArtList(r_ClassID)%>
	</div> 
	<!-- end content --> 
<!--#include file="C_Sidebar.asp" -->
</div> 
<!-- end page --> 
<!--#include file="C_Footer.asp" -->
</body> 
</html> 
<%APPPAGE="Default.asp"%>
<!--#include file="Comm.asp" -->
<!--#include file="API.asp" -->
<%
Page_Description=GetInf("Page_Description")
Page_Keywords=GetInf("Page_Keywords")
%>
<%PAGETITLE=GetInf("Page_Title")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"> 
<html xmlns="http://www.w3.org/1999/xhtml"> 
<!--#include file="C_Head.asp" -->
<body> 
<!--#include file="C_Header.asp" -->
<!-- start page --> 
<div id="page"> 
	<!-- start content --> 
	<div id="content">
		<%Call ReLoad()%>
		<!-- 输出最新文章 --> 
		<%Print_TopArt(GetInf("Num_DefaultArt"))%>
	</div> 
	<!-- end content --> 
<!--#include file="C_Sidebar.asp" -->
</div> 
<!-- end page --> 
<!--#include file="C_Footer.asp" -->
</body> 
</html> 
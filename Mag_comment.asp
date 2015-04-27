<%PCHK="ON"%>
<!--#include file="Comm.asp" -->
<!--#include file="API.asp" -->
<%
r_ID=Request.QueryString("ID")
AnitInjection(r_ID)
PAGETITLE="后台管理 "
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"> 
<html xmlns="http://www.w3.org/1999/xhtml"> 
<!--#include file="MC_Head.asp" -->
<body> 
<!--#include file="MC_Header.asp" -->
<!-- start page --> 
<div id="page"> 
	<!-- start content --> 
	<div id="content">
		<%If (Request.QueryString("c")="d") and (r_ID<>0) Then M_DelCom(r_ID) %>
		<!-- 输出评论列表 --> 
		<%M_Print_ComList()%>
	</div> 
	<!-- end content --> 
<!--#include file="MAG_Sidebar.asp" -->
</div> 
<!-- end page --> 
<!--#include file="MC_Footer.asp" -->
</body> 
</html> 
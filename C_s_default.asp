<!--#include file="Comm.asp" -->
<!--#include file="API.asp" -->
<%
PAGETITLE="sBlog"
Dim a_Title,a_Content,a_Class
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"> 
<html xmlns="http://www.w3.org/1999/xhtml"> 
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" /> 
<title><%=PAGETITLE%><%=GetInf("Page_Title")%></title> 
</head> 

<body style="background-color:#EAEAEA; font-size:12px;"> 

<!-- start page --> 
<div id="page"> 
	<!-- start content --> 
	<div id="contents">
		<div class="post">
			<h3>欢迎使用 sBlog Windows 管理工具</h3>
			您可以：
			<ul>
				<li>登陆您博客的后台进行管理</li>
				<li><a href="http://suomblog.cn/SuoMSoftware/SBlog_GUI.html" target="_blank">访问sBlog支持站点得到帮助</a></li>
				<li><a href="http://suomblog.cn/#About" target="_blank">联系sBlog作者</a></li>
			</ul>
			
		</div>
	</div> 
	<!-- end content --> 

</div> 
<!-- end page --> 

</body> 
</html> 
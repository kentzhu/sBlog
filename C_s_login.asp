<!--#include file="Comm.asp" -->
<!--#include file="API.asp" -->
<%
PAGETITLE="��½��̨"
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
			<form method="post" action="C_S_Article.asp" class="Sidebarform">
			<p>
			<label>�û�:</label>
			<input name="Username" value="" size="15" type="text" />
			<label>����:</label>
			<input name="Password" value="" size="15" type="password" />
			<input class="button" type="submit" value=" ��¼ " />
			</p>
			</form>
		</div>
	</div> 
	<!-- end content --> 

</div> 
<!-- end page --> 

</body> 
</html> 
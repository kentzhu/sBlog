<%PCHK="ON"%>
<!--#include file="Comm.asp" -->
<!--#include file="API.asp" -->
<%
PAGETITLE="��̨���� "
p1=Request.Form("Password1")
p2=Request.Form("Password2")
p3=Request.Form("Password3")
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
		<%If (p1<>"") and (p2<>"") and (p3<>"") Then Call M_ChgPD(p1,p2,p3) %>
		<div class="post">
			<h1 class="title">�޸�����</h1>
            <%Call NetChk()%>
			<form method="post" action="">
			�����룺<br />
			<input type="password" name="Password1" value="" /><br />
			�����룺<br />
			<input type="password" name="Password2" value="" /><br />
			�ظ����룺<br />
			<input type="password" name="Password3" value="" /><br />
			<input type="submit" value="�޸�����" />
			</form>
		</div>
	</div> 
	<!-- end content --> 
<!--#include file="MAG_Sidebar.asp" -->
</div> 
<!-- end page --> 
<!--#include file="MC_Footer.asp" -->
</body> 
</html> 
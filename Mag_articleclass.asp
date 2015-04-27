<%PCHK="ON"%>
<!--#include file="Comm.asp" -->
<!--#include file="API.asp" -->
<%
r_ID=Request.QueryString("ID")
c_ID=Request.Form("c_id")
AnitInjection(r_ID)
AnitInjection(c_ID)
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
		<%If (Request.QueryString("c")="s") Then Call M_SaveArtCla(c_ID,Request.Form("c_name")) %>
		<%If (Request.QueryString("c")="n") Then Call M_NewArtCla(Request.Form("c_name")) %>
		<%If (Request.QueryString("c")="d") and (r_ID<>0) Then Call M_DelArtCla(r_ID) %>
		<!-- 输出文章分类 --> 
		<%M_Print_ArtClaList()%>
		<%If (Request.QueryString("c")="e") and (r_ID<>0) Then Call M_PrintFormArtCla(r_ID) %>
		<form method="post" action="?c=n">
			新分类名称：<br />
			<input type="text" name="c_name" value="" /><br />
			<input type="submit" value="修改" />
		</form>
	</div> 
	<!-- end content --> 
<!--#include file="MAG_Sidebar.asp" -->
</div> 
<!-- end page --> 
<!--#include file="MC_Footer.asp" -->
</body> 
</html> 
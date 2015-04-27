<!--#include file="Comm.asp" -->
<!--#include file="API.asp" -->
<%
Dim STR_COMMENT
r_ID=Request.QueryString("ID")
AnitInjection(r_ID)

PAGETITLE=GetArticleTitle(r_ID)

Page_Keywords=PAGETITLE
Page_Description="文章:" & PAGETITLE

Call ReceiveComment()
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
		<!-- 输出文章 --> 
		<%Call Print_ArtView(r_ID)%>
        <!-- 输出前后文章链接 --> 
        <%Call Print_ArtUpNext(r_ID)%>
		<!-- 输出评论 --> 
		<%Call Print_CommentView(r_ID,False)%>
		<!-- 输出评论表单 -->
		<%Call Print_CommentBox()%>

	</div> 
	<!-- end content --> 
<!--#include file="C_Sidebar.asp" -->
</div> 
<!-- end page --> 
<!--#include file="C_Footer.asp" -->
</body> 
</html> 
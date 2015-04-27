<%PCHK="ON"%>
<!--#include file="Comm.asp" -->
<!--#include file="API.asp" -->
<%
PAGETITLE="后台管理 "
p1=Request.QueryString("Password1")
p2=Request.QueryString("Password2")
p3=Request.QueryString("Password3")
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
		<div class="post">
			<%If Request.Form("k")="c" Then Call SaveInf()%>
			<h1 class="title">基础设置</h1>
            <%Call NetChk()%>
			<form method="post" action="">
			<input type="hidden" name="k" value="c" />
            <h3>显示设置</h3>
			主标题：<br />
			<input type="text" name="Page_Title_1" value="<%=GetInf("Page_Title_1")%>" style="width:350px;" /><br />
			副标题：<br />
			<input type="text" name="Page_Title_2" value="<%=GetInf("Page_Title_2")%>" style="width:350px;" /><br />
			站点根目录：<br />
			<input type="text" name="Site_Root" value="<%=CNN_Site_Root%>" style="width:350px;" /><br />
			首页显示文章数：<br />
			<input type="text" name="Num_DefaultArt" value="<%=GetInf("Num_DefaultArt")%>" style="width:350px;" /><br />
			页面标题：<br />
			<input type="text" name="Page_Title" value="<%=GetInf("Page_Title")%>" style="width:350px;" /><br />
            <h3>SEO设置</h3>
			关键字：<br />
			<input type="text" name="Page_Keywords" value="<%=GetInf("Page_Keywords")%>" style="width:350px;" /><br />
			站点描述：<br />
			<input type="text" name="Page_Description" value="<%=GetInf("Page_Description")%>" style="width:350px;" /><br />

			<h3>功能设置</h3>
			开发评论（0关闭，1开放）：<br />
			<input type="text" name="Button_Comment" value="<%=GetInf("Button_Comment")%>" style="width:350px;" /><br />
            <h3>个人设置</h3>
            姓名：<br />
            <input type="text" name="Blogger_Name" value="<%=GetInf("Blogger_Name")%>" style="width:350px;" /><br />
            邮箱：<br />
            <input type="text" name="Blogger_Mail" value="<%=GetInf("Blogger_Mail")%>" style="width:350px;" /><br />
            简介（支持HTML代码）：<br />
            <textarea name="Blogger_Con" style="width:350px;"><%=GetInf("Blogger_Con")%></textarea><br />
            <h3>其他设置</h3>
            页面底部统计代码：<br />
            <textarea name="Footer_Code" style="width:350px;"><%=GetInf("Footer_Code")%></textarea>
			
            CSS模板路径[相对路径]：<br />
            <input type="text" name="Page_CSS" value="<%=GetInf("Page_CSS")%>" style="width:350px;" /><br />
			<input type="submit" value="完成" /><br />
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
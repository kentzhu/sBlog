<%PCHK="ON"%>
<!--#include file="Comm.asp" -->
<!--#include file="API.asp" -->
<%
r_ID=Request.QueryString("ID")
AnitInjection(r_ID)
PAGETITLE="后台管理 "
Dim a_Title,a_Content,a_Class
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"> 
<html xmlns="http://www.w3.org/1999/xhtml"> 
<!--#include file="MC_Head.asp" -->
<script type="text/javascript" charset="utf-8" src="Editor/kindeditor.js"></script>
<script type="text/javascript">
	KE.show({
	id : 'contentT',
	skinType: 'tinymce',
	cssPath : './index.css',
	items : [
		'source', 'preview', 'fullscreen', 'undo', 'redo', 'plainpaste', 'wordpaste', 'justifyleft', 'justifycenter', 'justifyright','justifyfull',
		 'insertorderedlist', 'insertunorderedlist', 'indent', 'outdent', 'subscript','superscript','-',
		 'title', 'fontname', 'fontsize', 'textcolor', 'bgcolor', 'bold','italic', 'underline','-',
		 'specialchar', 'emoticons', 'link', 'unlink', 'strikethrough', 'removeformat', 'selectall', 'image', 'flash', 'media', 'table', 'hr', 'about'
	]
	});
</script>

<body> 
<!--#include file="MC_Header.asp" -->
<!-- start page --> 
<div id="page"> 
	<!-- start content --> 
	<div id="content">
		<div class="post">

		<%If Request.QueryString("c")<>"d" Then%>
		
			<%Call M_DoUpdate()%>
			
			<%If Request.QueryString("c")="" Then%>
				<h1 class="title">撰写文章</h1>
				<%Call M_DoNew()%>
			<%End If%>
			<%If Request.QueryString("c")="e" Then%>
				<h1 class="title">编辑文章</h1>
				<%Call ReadArticle(r_ID)%>
			<%End If%>
			
            <%Call NetChk()%>
			
			<form method="post" action="">
			<input name="c_id" value="<%=r_id%>" type="hidden" />
			标题：<br />
			<%Call M_ArticleClass(a_Class)%><input type="text" name="Title" value="<%=a_Title%>" /><br />
			内容：<br />
			<textarea id="contentT" name="contentT" style="width:100%;height:400px;visibility:hidden;"><%=a_Content%></textarea>
			<input type="submit" value="提交" />
			</form>
		<%Else%>
			<%M_DelArt(r_id)%>
		<%End If%>
		</div>
		
	</div> 
	<!-- end content --> 
<!--#include file="MAG_Sidebar.asp" -->
</div> 
<!-- end page --> 
<!--#include file="MC_Footer.asp" -->
</body> 
</html> 
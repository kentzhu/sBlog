<%PCHK="ON"%>
<!--#include file="Comm.asp" -->
<!--#include file="API.asp" -->
<%
PAGETITLE="后台管理 "
Dim a_Title,a_Content,a_Class
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"> 
<html xmlns="http://www.w3.org/1999/xhtml"> 
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" /> 
<title><%=PAGETITLE%><%=GetInf("Page_Title")%></title> 
</head> 
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

<body style="background-color:#EAEAEA; font-size:12px;"> 

<!-- start page --> 
<div> 
	<!-- start content --> 
	<div>
		<div class="post">
			<%Call M_DoUpdate()%>
			<%If Request.QueryString("c")="" Then%>
				<%Call M_DoNew()%>
			<%End If%>
			<form method="post" action="">
			<input name="c_id" value="<%=r_id%>" type="hidden" />
			<%Call M_ArticleClass(a_Class)%>
			<input type="text" name="Title" value="<%=a_Title%>" size="40" />
			<input type="submit" value=" 发表 " /><br />
			<textarea id="contentT" name="contentT" style="width:100%;height:290px;visibility:hidden;"><%=a_Content%></textarea>
			
			</form>
		</div>
	</div> 
	<!-- end content --> 

</div> 
<!-- end page --> 

</body> 
</html> 
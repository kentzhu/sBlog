<%PCHK="ON"%>
<!--#include file="Comm.asp" -->
<!--#include file="API.asp" -->
<%
r_ClassID=Request.QueryString("ClassID")
AnitInjection(r_ClassID)
PAGETITLE="后台管理 " & GetArticleClassName(r_ClassID) & " "
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
		<!-- 输出文章列表 --> 
		<div class="post">
		<%If request.QueryString("c")="RDate" Then Call RDataBase()%>
		<%Call M_Print_SystemInfo()%>
		
        <h1 class="title">系统修复</h1>
        <table width="100%">
        <tbody>
            <tr>
                <td><a href="Mag_System.asp?c=RDate">数据库修复</a></td>
            </tr>
        </tbody>
        </table>
        <h1 class="title">技术帮助</h1>
        <table width="100%">
        <tbody>

            <tr>
                <td>邮箱：</td>
                <td>smjs.hi@163.com</td>
            </tr>
            <tr>
                <td>QQ：</td>
                <td>575658638</td>
            </tr>
            <tr>
                <td>sBlog客户中心：</td>
                <td><a href="http://suomblog.cn/sblog/Default.asp" target="_blank">http://suomblog.cn/sblog/Default.asp</a></td>
            </tr>
            <tr>
                <td>博客：</td>
                <td><a href="http://suomblog.cn/blog" target="_blank">http://suomblog.cn/blog</a></td>
            </tr>
        </tbody>
        </table>
		</div>

	</div> 
	<!-- end content --> 
<!--#include file="MAG_Sidebar.asp" -->
</div> 
<!-- end page --> 
<!--#include file="MC_Footer.asp" -->
</body> 
</html> 
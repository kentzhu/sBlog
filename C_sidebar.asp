	<!-- start sidebar --> 
	<div id="sidebar"> 
		<ul> 
			<li>
				<h2>导航:</h2>
				<ul>
					<li><a href="<%=CNN_Site_Root%>Default.asp">首页</a></li>
					<li><a href="<%=CNN_Site_Root%>Article.asp">文章</a></li>
					<li><a href="<%=CNN_Site_Root%>Search.asp">搜索</a></li>
                    <li><a href="<%=CNN_Site_Root%>rss.asp" target="_blank">RSS订阅</a></li>
					<%If session("user_name")<>"" Then%>
					<li><a href="<%=CNN_Site_Root%>MAG_Default.asp">管理</a></li>
					<li><a href="<%=CNN_Site_Root%>Default.asp?logout=on">注销</a></li>
					<%End If%>
				</ul>
			</li>
			<!-- 输出Blogger信息 -->
			<%If APPPAGE="Default.asp" Then Call Print_BloggerInfo()%>
			<!-- 输出搜索框 -->
			<%Call Print_SearchForm()%>
			<!-- 输出文章分类 -->
			<%Call Print_ArticleClass()%>
			<!-- 输出最新评论 -->
			<%If APPPAGE="Default.asp" Then Call Print_NewestComments()%>
            <!-- 随机输出文章 -->
            <%Call Print_ArtRnd()%>
			<!-- 输出推荐链接 -->
			<%If APPPAGE="Default.asp" Then Call Print_Link()%>
			<!-- 输出用户登录表单 -->
			<%If (session("user_name")="") and (APPPAGE="Default.asp") Then Call Print_SidebarLoginForm()%>

            <!-- 输出音乐播放器 -->
			<%'Call Print_MusicPlayer()%>

		</ul> 
  </div> 
  <!-- end sidebar --> 
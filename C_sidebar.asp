	<!-- start sidebar --> 
	<div id="sidebar"> 
		<ul> 
			<li>
				<h2>����:</h2>
				<ul>
					<li><a href="<%=CNN_Site_Root%>Default.asp">��ҳ</a></li>
					<li><a href="<%=CNN_Site_Root%>Article.asp">����</a></li>
					<li><a href="<%=CNN_Site_Root%>Search.asp">����</a></li>
                    <li><a href="<%=CNN_Site_Root%>rss.asp" target="_blank">RSS����</a></li>
					<%If session("user_name")<>"" Then%>
					<li><a href="<%=CNN_Site_Root%>MAG_Default.asp">����</a></li>
					<li><a href="<%=CNN_Site_Root%>Default.asp?logout=on">ע��</a></li>
					<%End If%>
				</ul>
			</li>
			<!-- ���Blogger��Ϣ -->
			<%If APPPAGE="Default.asp" Then Call Print_BloggerInfo()%>
			<!-- ��������� -->
			<%Call Print_SearchForm()%>
			<!-- ������·��� -->
			<%Call Print_ArticleClass()%>
			<!-- ����������� -->
			<%If APPPAGE="Default.asp" Then Call Print_NewestComments()%>
            <!-- ���������� -->
            <%Call Print_ArtRnd()%>
			<!-- ����Ƽ����� -->
			<%If APPPAGE="Default.asp" Then Call Print_Link()%>
			<!-- ����û���¼�� -->
			<%If (session("user_name")="") and (APPPAGE="Default.asp") Then Call Print_SidebarLoginForm()%>

            <!-- ������ֲ����� -->
			<%'Call Print_MusicPlayer()%>

		</ul> 
  </div> 
  <!-- end sidebar --> 
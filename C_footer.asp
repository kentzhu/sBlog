<div style="clear:both;"></div>
<div id="footer"> 
	<p>
	Power By <a href="http://suomblog.cn">sBlog v<%=GetInf("Sys_V")%></a>
	</p> 
	<%If session("user_name")<>"" Then Response.Write(GetInf("Footer_Code") & "<br />")%>	
    
    <!-- ���µĴ�����Է�ֹ��ʱ-->
	<script language="javascript"> 
    setInterval("Notimeout()","600000");
    function Notimeout()
    {
     var XmlHttp2 = new ActiveXObject("Msxml2.XMLHTTP")
     XmlHttp2.Open( "POST","NoTimeOut.asp", false );
     XmlHttp2.Send();    
    }
    </script>
	<br />
	
	����ʱ�䣺<%=GetLoadTime%>s
</div> 
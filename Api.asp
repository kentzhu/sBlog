<!--#include file="md5.asp"-->
<%




'----------------------------------------------------------------------------------------------------------+
'---------------------------------------------------------------- Function ���� ---------------------------+
'----------------------------------------------------------------------------------------------------------+




'��ȡ������Ϣ��ָ��Key���ƣ�����ֵ��
Function GetInf(Key)
	sqlCommand="Select * From [Inf] Where [Key]='"&Key&"'"
	Set rsGetInf=Server.CreateObject("ADODB.RecordSet")
	rsGetInf.Open sqlCommand,Conn,1,1
	If not rsGetInf.eof Then
		GetInf=rsGetInf("Value")
	Else
		GetInf="Error on Key["&Key&"]"
	End If
	rsGetInf.Close
	Set rsGetInf = Nothing
End Function

'��ȡ���·������ơ�ָ������ID�����ط�������
Function GetArticleClassName(id)
	If Not id=0 Then
		sqlCommand="Select * From [ArticleClass] Where [id]="&id
		Set rsArticleClassName=Server.CreateObject("ADODB.RecordSet")
		rsArticleClassName.Open sqlCommand,Conn,1,1
		If not rsArticleClassName.eof Then
			 GetArticleClassName=rsArticleClassName("Name")
		Else
			 GetArticleClassName="Error on ClassID["&id&"]"
		End If
		rsArticleClassName.Close
		Set rsArticleClassName = Nothing
	End If
End Function

'��ȡ���±��⡣ָ������ID���������±��⡣
Function GetArticleTitle(id)
	If Not id=0 Then
		sqlCommand="Select * From [Article] Where [id]="&id
		Set rsArticleTitle=Server.CreateObject("ADODB.RecordSet")
		rsArticleTitle.Open sqlCommand,Conn,1,1
		If not rsArticleTitle.eof Then
			 GetArticleTitle=rsArticleTitle("Title")
		Else
			 GetArticleTitle="Error on ArticleID["&id&"]"
		End If
		rsArticleTitle.Close
		Set rsArticleTitle = Nothing
	End If
End Function

'��ȡ����������ָ������ID��������Ӧ����������IDΪ-1���򷵻�������������
Function GetCount(ID)
	If ID=-1 Then
		sqlCommand="Select * From [Comment]"
	Else
		sqlCommand="Select * From [Comment] Where [Belong]="&id
	End If
	Set rsGetCount=Server.CreateObject("ADODB.RecordSet")
	rsGetCount.Open sqlCommand,Conn,1,1
	GetCount=rsGetCount.RecordCount
End Function

Function GetTCount(TName)
	sqlCommand="Select * From ["&TName&"]"
	Set rsGetCount=Server.CreateObject("ADODB.RecordSet")
	rsGetCount.Open sqlCommand,Conn,1,1
	GetTCount=rsGetCount.RecordCount
End Function

'��ȡ���¡�ָ������ID��������Ӧ����������IDΪ-1���򷵻�������������
Function GetArtCount(ID)
	If ID=-1 Then
		sqlCommand="Select * From [Aritcle]"
	Else
		sqlCommand="Select * From [Article] Where [Class]="&id
	End If
	Set rsArtGetCount=Server.CreateObject("ADODB.RecordSet")
	rsArtGetCount.Open sqlCommand,Conn,1,1
	GetArtCount=rsArtGetCount.RecordCount
End Function

'������֤����𰸡����ɴ𰸵�Session���������⡣
Function MakeQuestion()

	select case RndNumber(2,0)
	case 0
		a=RndNumber(50,0)
		b=RndNumber(50,0)
		c=a+b
		Session("QUESTION_NUM")=c
		MakeQuestion=a & " + " & b & "=?"
	case 1
		a=RndNumber(10,0)
		b=RndNumber(10,0)
		c=a*b
		Session("QUESTION_NUM")=c
		MakeQuestion=a & " * " & b & "=?"
	case 2
		a=RndNumber(100,0)
		b=RndNumber(100,0)
		c=a-b
		Session("QUESTION_NUM")=c
		MakeQuestion=a & " - " & b & "=?"
	end select
	
End Function

'��ע���⡣ָ���ַ�������鲻�����ֵģ���ʾ���󲢽�����
Sub AnitInjection(Str)
	If Not IsNumeric(Str) Then
		Response.Write("��Ǹ���������˷Ƿ�������������Դ�ɿ��ԡ�")
		Response.End()
	End If
End Sub

'��������������ֵ��Сֵ,���������
Function RndNumber(MaxNum,MinNum)
	Randomize 
	RndNumber=int((MaxNum-MinNum+1)*rnd+MinNum)
	RndNumber=RndNumber
End Function

'Gravatar ͷ��֧�ֺ���
Function MakeGravatarURL(mail)
	mail=trim(mail)
	If mail="" Then
		MakeGravatarURL = CNN_Site_Root & "images/guest.jpg"
	Else
		set h = new MD5
		md5str = h.hash(mail)
		defaultstr = server.urlEncode(CNN_Site_Root & "images/guest.jpg")
		MakeGravatarURL="http://www.gravatar.com/avatar.php?gravatar_id=" & md5str & "&default=" & defaultstr
	End If
End Function

'��������
Function return_RFC822_Date(byVal myDate, byVal TimeZone) 
	Dim myDay, myDays, myMonth, myYear 
	Dim myHours, myMinutes, mySeconds 
	myDate = CDate(myDate) 
	myDay = EnWeekDayName(myDate) 
	myDays = Right("00" & Day(myDate),2) 
	myMonth = EnMonthName(myDate) 
	myYear = Year(myDate) 
	myHours = Right("00" & Hour(myDate),2) 
	myMinutes = Right("00" & Minute(myDate),2) 
	mySeconds = Right("00" & Second(myDate),2) 
	
	return_RFC822_Date = myDay &", " & _ 
	myDays&" "& _ 
	myMonth&" "& _ 
	myYear&" "& _ 
	myHours&":"& _ 
	myMinutes&":"& _ 
	mySeconds&" "& _ 
	" " & TimeZone 
End Function 

'�·�ת��
Function EnWeekDayName(InputDate) 
	Dim Result 
	Select Case WeekDay(InputDate,1) 
		Case 1:Result="Sun" 
		Case 2:Result="Mon" 
		Case 3:Result="Tue" 
		Case 4:Result="Wed" 
		Case 5:Result="Thu" 
		Case 6:Result="Fri" 
		Case 7:Result="Sat" 
	End Select 
	EnWeekDayName = Result 
End Function 

'����ת��
Function EnMonthName(InputDate) 
	Dim Result 
	Select Case Month(InputDate) 
		Case 1:Result="Jan" 
		Case 2:Result="Feb" 
		Case 3:Result="Mar" 
		Case 4:Result="Apr" 
		Case 5:Result="May" 
		Case 6:Result="Jun" 
		Case 7:Result="Jul" 
		Case 8:Result="Aug" 
		Case 9:Result="Sep" 
		Case 10:Result="Oct" 
		Case 11:Result="Nov" 
		Case 12:Result="Dec" 
	End Select 
	EnMonthName = Result 
End Function 

'��ȥHTMLԪ��
Function nohtml(str) 
	dim re 
	Set re=new RegExp 
	re.IgnoreCase =true 
	re.Global=True 
	re.Pattern="<(.[^>]*)>" 
	str=re.replace(str,"") 
	nohtml=str 
	set re=nothing 
End Function

'�����ȡHTML
Function HTMLCutter(str, l)
    
    Dim sFIFO(100)
    
    p = 1
    str = Trim(str)
    maxlong = Len(str)
    
    FlagHTML = False
    longStr = 0
    
    For i = 1 To maxlong
        c = Mid(str, i, 1)
        If c = "<" Then
            If FlagHTML Then longStr = longStr + i - istart
            FlagHTML = True
            istart = i
        Else
            If FlagHTML Then
                If c = ">" Then
                    iend = i
                    tmpStr = Mid(str, istart + 1, iend - istart - 1)
                    
                    '���Ԫ��Ϊ��ʼ���ǽ���
                    '�ж��Ƿ���<XXX />�ı�ǩ
                    If Right(tmpStr, 1) <> "/" Then
                        If Left(tmpStr, 1) = "/" Then
                            '������ǩ
                            If sFIFO(p) = Right(tmpStr, Len(tmpStr) - 1) Then p = p - 1
                            
                            FlagHTML = False
                        Else
                            '��ʼ��ǩ
                            p = p + 1
                            
                            t = InStr(1, tmpStr, " ")
                            If t <> 0 Then tmpStr = Left(tmpStr, t - 1)
                            sFIFO(p) = tmpStr
                        End If
                    End If
                End If
            Else
                longStr = longStr + 1
                If longStr >= l Then Exit For
            End If
        End If
    Next
    
    fStr = Left(str, i)
    
    For j = p To 1 Step -1
        If sFIFO(j) <> "" Then endStr = endStr & " </" & sFIFO(j) & ">"
    Next
    
    HTMLCutter = fStr & endStr
End Function

'��ȡԶ��IP
Function GetResponseIP()
	r_IP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If r_IP="" Then r_IP = Request.ServerVariables("REMOTE_ADDR")
	GetResponseIP = r_IP
End Function

'��ȡ�ռ���Ϣ
Function ShowSpaceInfo(drvpath)
	dim fso,d,size,showsize
	set fso=server.createobject("scripting.filesystemobject")         
	drvpath=server.mappath(drvpath)                  
	set d=fso.getfolder(drvpath)         
	size=d.size
	showsize=size & " Byte" 
	if size>1024 then
	   size=(size\1024)
	   showsize=size & " KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & " MB"          
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & " GB"        
	end if   
	ShowSpaceInfo = showsize
End Function   

'��ȡ�ռ���Ϣ���׺���1
Function Drawbar(drvpath)
	dim fso,drvpathroot,d,size,totalsize,barsize
	set fso=server.createobject("scripting.filesystemobject")
	drvpathroot=server.mappath("pic")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	
	drvpath=server.mappath(drvpath)         
	set d=fso.getfolder(drvpath)
	size=d.size
	
	barsize=cint((size/totalsize)*400)
	Drawbar=barsize
End Function

'��ȡ�ռ���Ϣ���׺���2
Function Drawspecialbar()
	dim fso,drvpathroot,d,fc,f1,size,totalsize,barsize
	set fso=server.createobject("scripting.filesystemobject")
	drvpathroot=server.mappath("pic")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	
	set fc=d.files
	for each f1 in fc
		 size=size+f1.size
	next     
	
	barsize=cint((size/totalsize)*400)
	Drawspecialbar=barsize
End Function

'���������ؼ���(��ֹע��&XSS)
'
'�˺����İ�ȫ���д�����
'
Function DoSearchKeyWords(str)
    '��ֹ��վ����
    str=Server.HTMLEncode(str)
	'��ֹ������'��SQL�������Ч
    DoSearchKeyWords=replace(str,"'","''")
End Function

'������ҳ�����б�HTML����
Function MakeArtL(n)
    str=""
	sqlCommand="Select Top " & n & " * From [Article] Order By [Time] Desc"
	Set rsTopArt=Server.CreateObject("ADODB.RecordSet")
	rsTopArt.Open sqlCommand,Conn,1,1
	If Not rsTopArt.Eof Then
		While Not rsTopArt.Eof
			str=str& "<div class=""post""><h1 class=""title""><a href="""& CNN_Site_Root&"ArticleView.asp?id="&rsTopArt("id")&""" class=""title"">"& rsTopArt("Title")&"</a></h1><p class=""byline""><small> �� "&rsTopArt("By")&" ������ "& rsTopArt("Time")&" | ���ࣺ<a href="""&CNN_Site_Root&"Article.asp?ClassID="&rsTopArt("Class")&""">"& GetArticleClassName(rsTopArt("Class"))&"</a></small></p><div class=""entry""><p>"& HTMLCutter(rsTopArt("Content"),300)&" ... <a href="""&CNN_Site_Root&"ArticleView.asp?id="&rsTopArt("id")& """ class=""more"">�鿴ȫ��</a></p></div><p class=""meta""><a href="""&CNN_Site_Root& "ArticleView.asp?id="&rsTopArt("id")&""" class=""more"">�̶�����</a>&nbsp;&nbsp;&nbsp;<a href="""& CNN_Site_Root&"ArticleView.asp?id="&rsTopArt("id")& """ class=""comments"">���� ("&GetCount(rsTopArt("id"))&")</a></p></div>"
			'HTMLCutter(rsTopArt("Content"),300)
			rsTopArt.movenext
		WEnd
	Else
		str="��û������"
	End If
	rsTopArt.Close
	Set rsTopArt = Nothing
    MakeArtL=str
End Function

'
Function GetLoadTime()
	t=Timer-TimeStar
	If t<1 Then
		GetLoadTime="0" & t
	Else
		GetLoadTime=t
	End If
End Function

Function ChkBrowserIsIE()
	set cbii=server.CreateObject("MSWC.BrowserType")
	ChkBrowser=Flase
	If cbii.Browser="IE" Then ChkBrowserIsIE=True
	Response.Write(cbii.Browser)
End Function

'----------------------------------------------------------------------------------------------------------+
'---------------------------------------------------------------- Sub ���� --------------------------------+
'----------------------------------------------------------------------------------------------------------+




'--------------------------------+
'------------------��ӡҳ�沿��--+
'--------------------------------+


'����������
Sub Print_SearchList()
    str=request.Form("s")
    If (len(str)<>0) or (len(str)>19) then
		str=DoSearchKeyWords(str)
		sqlCommand="Select * From [Article] Where [Title] like '%" & str & "%' Order By [Time] Desc"
		'response.Write(sqlCommand)
		Set rsSL=Server.CreateObject("ADODB.RecordSet")
		rsSL.Open sqlCommand,Conn,1,1
		If not rsSL.eof then
			%><div class="post"><h1 class="title">�ؼ���"<%=str%>"���������</h1><table width="100%"><tbody><tr><th>����</th><th width="75px">����</th></tr><%
			While Not rsSL.Eof
				%><tr><td><a href="<%=CNN_Site_Root%>ArticleView.asp?id=<%=rsSL("id")%>"><%=rsSL("Title")%></a></td><td><%=GetArticleClassName(rsSL("Class"))%></td></tr><%
				rsSL.movenext
			WEnd
			%></tbody></table></div><%
		Else
			%><div class="post"><h1 class="title">�ؼ���"<%=str%>"���������</h1><p class="meta">�Բ���ϵͳδ�ҵ������ҵ���Ϣ��</p></div><%
		End If
		rsSL.Close
		Set rsSL = Nothing
	Else
	    %><div class="post"><h1 class="title">��������</h1></div><%
	End If
	    %><div class="post"><p class="meta">�����벻����12���ַ��Ĺؼ��ʣ�ϵͳ����֮�����±�����м�����</p><form method="post" action="Search.asp">�ؼ��֣�<input name="s" type="text" size="12" maxlength="12" value="<%=str%>" style="width:200px;" />
<input type="submit" value="����" style="width:75px;" /></form></div><%
	
End Sub

'�����ҳ���¡�ָ��nΪ��ʾ��������
Sub Print_TopArt(n)
    str=GetInf("Site_MainPage")
	If str="" Then Call Reload()
	Response.Write(str)
End Sub

'����������ġ�ָ��ID
Sub Print_ArtView(id)
	sqlCommand="Select * From [Article] Where [id]="&id
	Set rsArtView=Server.CreateObject("ADODB.RecordSet")
	rsArtView.Open sqlCommand,Conn,1,1
	If Not rsArtView.Eof Then
		%><div class="post"><h1 class="title"><a href="<%=CNN_Site_Root%>ArticleView.asp?id=<%=rsArtView("id")%>" class="title"><%=rsArtView("Title")%></a></h1><p class="byline"><small> �� <%=rsArtView("By")%> ������ <%=rsArtView("Time")%> | ���ࣺ<a href="<%=CNN_Site_Root%>Article.asp?ClassID=<%=rsArtView("Class")%>"><%=GetArticleClassName(rsArtView("Class"))%></a></small></p><div class="entry"><p><%=rsArtView("Content")%></p></div></div><div style="clear:bothl"></div><%
		rsArtView.movenext
	End If
	rsArtView.Close
	Set rsArtView = Nothing
End Sub

'��������б�ָ��IDΪ����ID��IDΪ0����ȫ����
Sub Print_ArtList(id)
	If id=0 Then
		sqlCommand="Select * From [Article] Order By [Time] Desc"
	Else
		sqlCommand="Select * From [Article] Where [Class]=" & id & " Order By [Time] Desc"
	End If
	Set rsArtList=Server.CreateObject("ADODB.RecordSet")
	rsArtList.Open sqlCommand,Conn,1,1
	%><div class="post"><h1 class="title">�����б� <%=GetArticleClassName(id)%></h1><table width="100%"><tbody><tr><th>����</th><th width="75px">����</th></tr><%
	While Not rsArtList.Eof
		%><tr><td><a href="<%=CNN_Site_Root%>ArticleView.asp?id=<%=rsArtList("id")%>"><%=rsArtList("Title")%></a></td><td><%=GetArticleClassName(rsArtList("Class"))%></td></tr><%
		rsArtList.movenext
	WEnd
	%></tbody></table></div><%
	rsArtList.Close
	Set rsArtList = Nothing
End Sub

'������·��ࣨ�������
Sub Print_ArticleClass()
	sqlCommand="Select * From [ArticleClass]"
	Set rsArticleClass=Server.CreateObject("ADODB.RecordSet")
	rsArticleClass.Open sqlCommand,Conn,1,1
	Response.Write("<li><h2>���·���:</h2><ul>")
	While Not rsArticleClass.Eof
		%><li><a href="<%=CNN_Site_Root%>Article.asp?ClassID=<%=rsArticleClass("id")%>"><%=rsArticleClass("Name")%></a> [<%=GetArtCount(rsArticleClass("id"))%>]</li><%
		rsArticleClass.movenext
	WEnd
	Response.Write("</ul></li>")
	rsArticleClass.Close
	Set rsArticleClass = Nothing
End Sub

'����Ƽ����ӣ��������
Sub Print_Link()
	sqlCommand="Select * From [Link]"
	Set rsLink=Server.CreateObject("ADODB.RecordSet")
	rsLink.Open sqlCommand,Conn,1,1
	Response.Write("<li><h2>�Ƽ�����:</h2><ul>")
	While Not rsLink.Eof
		%><li><a href="<%=rsLink("Url")%>" title="<%=rsLink("Remarks")%>" target="_blank"><%=rsLink("Name")%></a></li><%
		rsLink.movenext
	WEnd
	Response.Write("</ul></li>")
	rsLink.Close
	Set rsLink = Nothing
End Sub

'���������
Sub Print_SearchForm()
	%><li><h2>��������:</h2><form method="post" action="<%=CNN_Site_Root%>Search.asp" class="Sidebarform"><p>�ؼ���(12������)��<br /><input name="s" value="" maxlength="12" size="12" type="text" /><input class="button" type="submit" value="����" /></p></form></li><%
End Sub

'����û���½�����������
Sub Print_SidebarLoginForm()
	%><li><h2>�û���¼:</h2><form method="post" action="<%=CNN_Site_Root%>MAG_Default.asp" class="Sidebarform"><p>�û���<br /><input name="Username" value="" size="15" type="text" /><br />����<br /><input name="Password" value="" size="15" type="password" /><br /><input class="button" type="submit" value="��¼" /></p></form></li><%
End Sub

'������۱�
Sub Print_CommentBox()
	If GetInf("Button_Comment")="1" Then
%>

    
	<form action="" method="post">

		<%=STR_COMMENT%><br />
        <%if session("user_Name")="" then%>
        
		<script language="javascript"> 
        window.setTimeout("getf()","1000");
        setInterval("getf()","600000");
        function getf()
        {
         var XmlHttp = new ActiveXObject("Msxml2.XMLHTTP")
         XmlHttp.Open( "POST","getquestion.asp", false );
         XmlHttp.Send();
           if (XmlHttp.status == 200) document.getElementById("que").innerHTML = XmlHttp.responseText;
        
        }
        </script>
        
        
		����(�ɲ���):<br /><input name="c_Name" type="text" /><br /><br />
		����(�ɲ����ֻ�в����ܿ���):<br /><input name="c_Mail" type="text" /><br /><br />
		����(����):<br /><textarea name="c_Content"></textarea><br />
		��֤(����):<a onclick="getf();">��һ������</a><br /><span id="que"><%=MakeQuestion()%></span><br /><input name="c_Question" type="text" />
		<input type="submit" value="�ύ" />
        
        <%else%>
        
		����:<%=session("user_name")%><br /><br />
		����(����):<br /><textarea name="c_Content"></textarea><br />
		<input type="submit" value="�ύ" />
        
        <%end if%>
        

	</form>
<%
	Else
		%><div class="meta">���۹����ѹر�</div><%
	End If
End Sub

'��������б�ָ������ID�������Ӧ���ۡ�
Sub Print_CommentView(id,all)
	If all Then
		sqlCommand="Select * From [Comment] Where [Belong]=" & id & " Order By [Time] Desc"
	Else
		sqlCommand="Select Top 10 * From [Comment] Where [Belong]=" & id & " Order By [Time] Desc"
	End If
	Set rsCommentView=Server.CreateObject("ADODB.RecordSet")
	rsCommentView.Open sqlCommand,Conn,1,1
	If rsCommentView.Eof Then
		'No Comment
		Response.Write("<div class=""meta"">��û������</div>")
	Else
		While Not rsCommentView.Eof
			%>
			<div class="meta" style="height:90px;">
            	<img style="float:right; clear:both; padding-top:10px; padding-left:10px;" src="<%=MakeGravatarURL(rsCommentView("EMail"))%>" width="64" height="64" /><br />
				<strong><%=Server.HTMLEncode(" " & rsCommentView("Name"))%></strong> �� <%=rsCommentView("Time")%> ʱ˵:<br />
				<%=Server.HTMLEncode(" " & rsCommentView("Content"))%><br />
            </div>
			<%
			rsCommentView.Movenext
		WEnd
	End If
	If not all Then
		If GetCount(id)>10 Then
			%>
			<div class="meta">
            	����Ϊ�����µ����10�����ۣ�<a href="CommentAll.asp?id=<%=id%>" title="�鿴��<%=pagetitle%>������������" target="_blank">����˴��鿴��������</a>��
            </div>
			<%
		End If
	End If
End Sub

'����������ۣ��������
Sub Print_NewestComments()
	%><li><h2>����</h2><ul><%
	sqlCommand="Select Top 5 * From [Comment] Order By [Time] Desc"
	Set rsCommentView=Server.CreateObject("ADODB.RecordSet")
	rsCommentView.Open sqlCommand,Conn,1,1
	If rsCommentView.Eof Then
		Response.Write("��û������")
	Else
		While Not rsCommentView.Eof
			Response.Write("<li><a href=""" & CNN_Site_Root & "ArticleView.asp?id=" & rsCommentView("Belong") & """>" & rsCommentView("Name") & "��" & Server.HTMLEncode(" " & rsCommentView("Content")) & "</a></li>")
			rsCommentView.Movenext
		WEnd
	End If
	rsCommentView.Close
	Set rsCommentView=Nothing
	%></ul></li><%
End Sub

'���ǰ�����������
Sub Print_ArtUpNext(id)
    response.Write("<div class=""meta"">")
	
	sqlCommand="Select * From [Article] Where [id]>" & id & " Order By [id]"
	Set rsAUN=Server.CreateObject("ADODB.RecordSet")
	rsAUN.Open sqlCommand,Conn,1,1
	If not rsAUN.Eof Then
		%>��һƪ:<a href="<%=CNN_Site_Root%>ArticleView.asp?id=<%=rsAUN("id")%>" title="<%=rsAUN("Title")%>"><%=rsAUN("Title")%></a><br /><%
	End If
	rsAUN.close
	
	sqlCommand="Select * From [Article] Where [id]<" & id & " Order By [id] desc"
	Set rsAUN=Server.CreateObject("ADODB.RecordSet")
	rsAUN.Open sqlCommand,Conn,1,1
	If not rsAUN.Eof Then
		%>��һƪ:<a href="<%=CNN_Site_Root%>ArticleView.asp?id=<%=rsAUN("id")%>" title="<%=rsAUN("Title")%>"><%=rsAUN("Title")%></a><%
	End If
	rsAUN.close
	set rsAUN = Nothing
	response.Write("</div>")
End Sub

'��������������
Sub Print_ArtRnd()
	%><li><h2>����Ƽ�</h2><ul><%
	Randomize 
	sqlCommand="Select Top 5 * From [Article] Order By rnd(-([id]+" & rnd() & "))"
	Set rsArtRnd=Server.CreateObject("ADODB.RecordSet")
	rsArtRnd.Open sqlCommand,Conn,1,1
	while not rsArtRnd.Eof
		%><li><a href="<%=CNN_Site_Root%>ArticleView.asp?id=<%=rsArtRnd("id")%>" title="<%=rsArtRnd("Title")%>"><%=rsArtRnd("Title")%></a></li><%
        rsArtRnd.movenext
	wend
	rsArtRnd.close
	set ArtRnd = Nothing
	%></ul></li><%
End Sub

'�������������Ϣ���������
Sub Print_BloggerInfo()
	%>
    <li>
        <h2>������</h2>
        <ul>
            <li>����:<%=GetInf("Blogger_Name")%></li>
            <li>����:<%=GetInf("Blogger_Mail")%></li>
            <li>���:<%=GetInf("Blogger_Con")%></li>
        </ul>
    </li>
	<%
End Sub

'��������б�[�ֻ�ҳ��]��ָ������ID����IDΪ0���������¡�
Sub Print_M_ArtList(id)
	If id=0 Then
		sqlCommand="Select Top 5 * From [Article] Order By [Time] Desc"
	Else
		sqlCommand="Select Top 5 * From [Article] Where [Class]=" & id & " Order By [Time] Desc"
	End If
	Set rsArtList=Server.CreateObject("ADODB.RecordSet")
	rsArtList.Open sqlCommand,Conn,1,1
	While Not rsArtList.Eof
		%><a href="M_ArtV.asp?id=<%=rsArtList("id")%>"><%=rsArtList("Title")%></a><%
		rsArtList.movenext
	WEnd
	rsArtList.Close
	Set rsArtList = Nothing
End Sub

'������ֲ��ţ��������
Sub Print_MusicPlayer()
%>
    <li>
    <h2>���ֲ���:</h2>
    
		<script type="text/javascript" src="swfobject.js"></script>
                
        <div id="flashPlayer">
          �����������֧��Flash
        </div>
        
        <script type="text/javascript">
           var so = new SWFObject("playerMultipleList.swf", "mymovie", "240", "200", "7", "#FFFFFF");  
           so.addVariable("autoPlay","yes");
           so.addVariable("playlistPath","playlist.asp");
           so.write("flashPlayer");
        </script>
    </li>
<%
End Sub

'����״�����
Sub NetChk()
%>
<p>
	<script language="javascript"> 
    setInterval("NetChk()","60000");
    function NetChk()
    {
     var XmlHttp2 = new ActiveXObject("Msxml2.XMLHTTP")
     XmlHttp2.Open( "POST","NoTimeOut.asp", false );
     XmlHttp2.Send();
     if (XmlHttp2.status == 200) document.getElementById("NetChk").innerHTML = "";
     if (XmlHttp2.status != 200) document.getElementById("NetChk").innerHTML = "�����޷����ӵ�Զ�̷��������뱣�浱ǰ���������ء�";
    }
    </script>
<h3><span id="NetChk" style="color:#F00;"></span></h3>
</p>
<%
End Sub

'--------------------------------+
'------------------ִ�й��ܲ���--+
'--------------------------------+

'����վ�����ݼ�����
Sub ReLoad()
	Call SaveInfKey("Site_MainPage",MakeArtL(GetInf("Num_DefaultArt")))
End Sub


'��½���
Sub LoginChk()
	Username=Request.Form("Username")
	Password=Request.Form("Password")
	
	If (Request.Form("Username")<>"") and (Request.Form("Password")<>"") Then
		sqlCommand="Select * From [User] Where [username]='"&Username&"'"
		Set rsLoginChk=Server.CreateObject("ADODB.RecordSet")
		rsLoginChk.Open sqlCommand,Conn,1,1
		If Not rsLoginChk.eof Then
			If rsLoginChk("Password")=Password Then
				Session("user_id")=rsLoginChk("id")
				Session("user_Name")=rsLoginChk("name")
				Session("user_username")=rsLoginChk("username")
				Session("user_birthday")=rsLoginChk("birthday")
				Session("user_Email")=rsLoginChk("Email")
				Session("user_QICQ")=rsLoginChk("QICQ")
				Session("user_Group")=rsLoginChk("Group")
				Session("Error_Login")=""
			End If
		End If
		rsLoginChk.Close
		Set rsLoginChk=Nothing
	End If
End Sub

'ע����¼
Sub Logout()
	Session("user_id")=""
	Session("user_Name")=""
	Session("user_username")=""
	Session("user_birthday")=""
	Session("user_Email")=""
	Session("user_QICQ")=""
	Session("user_Group")=""
End Sub

'�����µ�a_class,a_Title,a_Content
Sub ReadArticle(id)
	sqlCommand="Select * From [Article] Where [id]="&id
	Set rsArticle=Server.CreateObject("ADODB.RecordSet")
	rsArticle.Open sqlCommand,Conn,1,1
	If Not rsArticle.Eof Then
		a_class=rsArticle("class")
		a_Title=rsArticle("Title")
		a_Content=rsArticle("Content")
		rsArticle.movenext
	End If
	rsArticle.Close
	Set rsArticle = Nothing
End Sub

'�������ۣ�����֤��
Sub ReceiveComment()

	Comm_IP=GetResponseIP()

	If GetInf("Button_Comment")="1" Then
		Dim Comm_IP
		
		Comm_Question=Trim(Request.Form("c_Question"))
		Comm_Name=Trim(Request.Form("c_Name"))
		Comm_Mail=Trim(Request.Form("c_Mail"))
		Comm_Content=Trim(Request.Form("c_Content"))

			  if (session("user_name")="") then
			  'δ��¼�û�
			  	  If Trim(Session("QUESTION_NUM"))<>"" Then
					  If (Comm_Question&" ")=(Session("QUESTION_NUM")&" ") Then
					  
	
						  If Comm_Name="" Then Comm_Name=Comm_IP
						  
						  sqlCommand="Select * From [Comment]"
						  Set rsComment=Server.CreateObject("ADODB.RecordSet")
						  rsComment.Open sqlCommand,Conn,1,3
						  rsComment.Addnew
							  rsComment("Name")="[�ο�]" & Comm_Name
							  rsComment("Time")=Now
							  rsComment("EMail")=Comm_Mail
							  rsComment("Content")=Comm_Content
							  rsComment("IP")=Comm_IP
							  rsComment("Belong")=r_ID
						  rsComment.Update
						  rsComment.Close
						  Set rsComment=Nothing
						  
					  Else
						  STR_COMMENT="<strong>��֤�����</strong>"
					  End If
				  End If
			  else
			  '��¼�û�
				  If  Comm_Content<>"" Then
					  sqlCommand="Select * From [Comment]"
					  Set rsComment=Server.CreateObject("ADODB.RecordSet")
					  rsComment.Open sqlCommand,Conn,1,3
					  rsComment.Addnew
						  rsComment("Name")="[�û�]" & session("user_name")
						  rsComment("Time")=Now
						  rsComment("EMail")=session("user_Email")
						  rsComment("Content")=Comm_Content
						  rsComment("IP")=Comm_IP
						  rsComment("Belong")=r_ID
					  rsComment.Update
					  rsComment.Close
					  Set rsComment=Nothing
				  End If
			  end if

	End If
	
	Session("QUESTION_NUM")=""
	
End Sub


'------
'------------
'--------------------------
'----------------------------------------------------------------          ��Ҫ���¼   ��
'--------------------------------------------------------------------------------------------------
'----------------------------------------------------------------          Ҫ���¼     ��
'--------------------------
'------------
'------

'��������б�[����]
Sub M_Print_ArtList(id)
	If id=0 Then
		sqlCommand="Select * From [Article] Order By [Time] Desc"
	Else
		sqlCommand="Select * From [Article] Where [Class]=" & id & " Order By [Time] Desc"
	End If
	Set rsArtList=Server.CreateObject("ADODB.RecordSet")
	rsArtList.Open sqlCommand,Conn,1,1
	%><div class="post"><h1 class="title">�����б� <%=GetArticleClassName(id)%></h1><%Call NetChk()%><table width="100%"><tbody><tr><th>����</th><th width="58px">�༭</th></tr><%
	While Not rsArtList.Eof
		%><tr><td><a href="ArticleView.asp?id=<%=rsArtList("id")%>" target="_blank"><%=rsArtList("Title")%></a></td><td><a href="MAG_Article.asp?c=e&id=<%=rsArtList("id")%>">�༭</a> <a href="MAG_Article.asp?c=d&id=<%=rsArtList("id")%>">ɾ��</a></td></tr><%
		rsArtList.movenext
	WEnd
	%></tbody></table></div><%
	rsArtList.Close
	Set rsArtList = Nothing
End Sub

'���ϵͳ��Ϣ
Sub M_Print_SystemInfo()
%>
        <h1 class="title">ϵͳ��Ϣ</h1>
        <table width="100%">
        <tbody>
            <tr>
                <th width="100px">��Ŀ</th>
                <th>ֵ</th>
            </tr>
            <tr>
                <td>����汾</td>
                <td><%=ChkVersion()%></td>
            </tr>
            <tr>
                <td>��ǰ����</td>
                <td><%=Application("ActiveUsers")%>��</td>
            </tr>
            <tr>
                <td>�ռ�ռ��</td>
                <td><%=ShowSpaceInfo(".\")%></td>
            </tr>
            <tr>
                <td>��������</td>
                <td><%=GetTCount("Article")%></td>
            </tr>
			<tr>
                <td>���·�������</td>
                <td><%=GetTCount("ArticleClass")%></td>
            </tr>
			<tr>
                <td>�ÿ���������</td>
                <td><%=GetTCount("Comment")%></td>
            </tr>
        </tbody>
        </table>
<%
End Sub

'��������б�[����]
Sub M_Print_ComList()
	sqlCommand="Select * From [Comment] Order By [Time] Desc"
	Set rsComList=Server.CreateObject("ADODB.RecordSet")
	rsComList.Open sqlCommand,Conn,1,1
	%>
	
	<div class="post">
	<h1 class="title">�����б�</h1>
	<%Call NetChk()%>
	<table width="100%">
	<tbody>
	<tr>
	<th width="30px">ID</th>
	<th width="50px">����</th>
	<th width="100px">IP</th>
	<th width="30px">����</th>
	</tr>
	<%
	While Not rsComList.Eof
	%>
		<tr>
		<td><%=rsComList("id")%></td>
		<td><%=Server.HTMLEncode(rsComList("Name"))%></td>
		<td><%=rsComList("IP")%></td>
		<td><a href="?c=d&id=<%=rsComList("id")%>">ɾ��</a></td>
		</tr>
		<tr>
		<td></td><td colspan="3"><%=Server.HTMLEncode(rsComList("Content"))%></td>
		</tr>
		<tr>
		<td colspan="4"></td>
		</tr>
		<%
		rsComList.movenext
	WEnd
	%>
	</tbody>
	</table>
	</div><%
	rsComList.Close
	Set rsComList = Nothing
End Sub

'������·����б�[����]
Sub M_Print_ArtClaList()
	sqlCommand="Select * From [ArticleClass]"
	Set rsComList=Server.CreateObject("ADODB.RecordSet")
	rsComList.Open sqlCommand,Conn,1,1
	%><div class="post"><h1 class="title">���·���</h1><%Call NetChk()%><table width="100%"><tbody><tr><th>��������</th><th width="60px">�༭</th></tr><%
	While Not rsComList.Eof
		%><tr><td><%=rsComList("Name")%></td><td><a href="?c=e&id=<%=rsComList("id")%>">�༭</a> <a href="?c=d&id=<%=rsComList("id")%>">ɾ��</a></td></tr><%
		rsComList.movenext
	WEnd
	%></tbody></table></div><%
	rsComList.Close
	Set rsComList = Nothing
End Sub

'����Ƽ������б�[����]
Sub M_Print_Link()
	sqlCommand="Select * From [Link]"
	Set rsComList=Server.CreateObject("ADODB.RecordSet")
	rsComList.Open sqlCommand,Conn,1,1
	%><div class="post"><h1 class="title">�Ƽ�����</h1><%Call NetChk()%><table width="100%"><tbody><tr><th>��������</th><th width="60px">�༭</th></tr><%
	While Not rsComList.Eof
		%><tr><td><%=rsComList("Name")%></td><td><a href="?c=e&id=<%=rsComList("id")%>">�༭</a> <a href="?c=d&id=<%=rsComList("id")%>">ɾ��</a></td></tr><%
		rsComList.movenext
	WEnd
	%></tbody></table></div><%
	rsComList.Close
	Set rsComList = Nothing
End Sub

'��������
Sub M_DoUpdate()
	If (Request.Form("c_id")<>"") and (Request.Form("Title")<>"") and (Request.Form("ContentT")<>"") Then
		sqlCommand="Select * From [Article] Where [id]=" & Request.Form("c_id")
		Set rsDoUpdate=Server.CreateObject("ADODB.RecordSet")
		rsDoUpdate.Open sqlCommand,Conn,1,3
		If Not rsDoUpdate.Eof Then
			rsDoUpdate("class")=Request.Form("classT")
			rsDoUpdate("Title")=Request.Form("Title")
			rsDoUpdate("Content")=Request.Form("ContentT")
			rsDoUpdate("By")=session("user_Name")
			rsDoUpdate.UpDate
			%><p class="meta">���¸��³ɹ���</p><%
		End If
		rsDoUpdate.Close
		Set rsDoUpdate = Nothing
		Reload()
	End If
End Sub

'�������
Sub M_DoNew()
	If (Request.Form("c_id")="") and (Request.Form("Title")<>"") and (Request.Form("ContentT")<>"") Then
		sqlCommand="Select * From [Article]"
		Set rsDoUpdate=Server.CreateObject("ADODB.RecordSet")
		rsDoUpdate.Open sqlCommand,Conn,1,3
		rsDoUpdate.addnew
		rsDoUpdate("class")=Request.Form("classT")
		rsDoUpdate("Title")=Request.Form("Title")
		rsDoUpdate("Content")=Request.Form("ContentT")
		rsDoUpdate("By")=session("user_name")
		rsDoUpdate("Time")=now
		r_id=rsDoUpdate("id")
		rsDoUpdate.UpDate
		%><p class="meta">������ӳɹ���</p><%
		rsDoUpdate.Close
		Set rsDoUpdate = Nothing
		Reload()
	End If
End Sub

'ɾ������
Sub M_DelArt(id)
	sqlCommand="DELETE FROM [Article] WHERE [ID]=" & id
	Set rsDoSQL = conn.Execute(sqlCommand)
	Set rsDoSQL = Nothing
	Reload()
	%><p class="meta">ָ��������ɾ��</p><%
End Sub

'ɾ������
Sub M_DelLink(id)
	sqlCommand="DELETE FROM [Link] WHERE [ID]=" & id
	Set rsDoSQL = conn.Execute(sqlCommand)
	Set rsDoSQL = Nothing
	%><p class="meta">ָ��������ɾ��</p><%
End Sub

'ɾ������
Sub M_DelArtCla(id)
    If GetArtCount(id)>0 Then
		%><p class="meta">�Բ��𣬸÷����»������£��޷�ɾ���˷��ࡣ</p><%
	Else
		sqlCommand="DELETE FROM [ArticleClass] WHERE [ID]=" & id
		Set rsDoSQL = conn.Execute(sqlCommand)
		Set rsDoSQL = Nothing
		%><p class="meta">ָ�����·�����ɾ��</p><%
	End If
End Sub

'ɾ������
Sub M_DelCom(id)
	sqlCommand="DELETE FROM [Comment] WHERE [ID]=" & id
	Set rsDoSQL = conn.Execute(sqlCommand)
	Set rsDoSQL = Nothing
	%><p class="meta">ָ��������ɾ��</p><%
End Sub

'��������
Sub M_ChgPD(p1,p2,p3)
	If p2=p3 Then
		sqlCommand="Select * From [User] Where [id]=" & session("user_id")
		Set rsChgPD=Server.CreateObject("ADODB.RecordSet")
		rsChgPD.Open sqlCommand,Conn,1,3
		If p1=rsChgPD("password") Then
			rsChgPD("password")=p2
			rsChgPD.UpDate
			%><p class="meta">�����޸ĳɹ�</p><%
		Else
			%><p class="meta">ԭ�������</p><%
		End If
		rsChgPD.Close
		Set rsChgPD = Nothing
	Else
		%><p class="meta">�������벻һ��</p><%
	End If
End Sub

'�޸����·���
Sub M_SaveArtCla(c_id,c_name)
	sqlCommand="Select * From [ArticleClass] Where [id]=" & c_id
	Set rsSaveArtCla=Server.CreateObject("ADODB.RecordSet")
	rsSaveArtCla.Open sqlCommand,Conn,1,3
	If Not rsSaveArtCla.Eof Then
		rsSaveArtCla("name")=c_name
		rsSaveArtCla.UpDate
		%><p class="meta">���·����޸ĳɹ�</p><%
	End If
	rsSaveArtCla.Close
	Set rsSaveArtCla = Nothing
End Sub

'�޸��Ƽ�����
Sub M_SaveLink(c_id,c_name,c_url)
	sqlCommand="Select * From [Link] Where [id]=" & c_id
	Set rsSaveLink=Server.CreateObject("ADODB.RecordSet")
	rsSaveLink.Open sqlCommand,Conn,1,3
	If Not rsSaveLink.Eof Then
		rsSaveLink("name")=c_name
		rsSaveLink("url")=c_url
		rsSaveLink.UpDate
		%><p class="meta">�Ƽ������޸ĳɹ�</p><%
	End If
	rsSaveLink.Close
	Set rsSaveLink = Nothing
End Sub

'������·���
Sub M_NewArtCla(c_name)
	sqlCommand="Select * From [ArticleClass]"
	Set rsNewArtCla=Server.CreateObject("ADODB.RecordSet")
	rsNewArtCla.Open sqlCommand,Conn,1,3
	rsNewArtCla.Addnew
	rsNewArtCla("name")=c_name
	rsNewArtCla.UpDate
	%><p class="meta">���·�����ӳɹ�</p><%
	rsNewArtCla.Close
	Set rsNewArtCla = Nothing
End Sub

'����Ƽ�����
Sub M_NewLink(c_name,c_url)
	sqlCommand="Select * From [Link]"
	Set rsNewLink=Server.CreateObject("ADODB.RecordSet")
	rsNewLink.Open sqlCommand,Conn,1,3
	rsNewLink.Addnew
	rsNewLink("name")=c_name
	rsNewLink("url")=c_url
	rsNewLink.UpDate
	%><p class="meta">�Ƽ�������ӳɹ�</p><%
	rsNewLink.Close
	Set rsNewLink = Nothing
End Sub

'������·����
Sub M_PrintFormArtCla(id)
	sqlCommand="Select * From [ArticleClass] Where [id]=" & id
	Set rsPrintFormArtCla=Server.CreateObject("ADODB.RecordSet")
	rsPrintFormArtCla.Open sqlCommand,Conn,1,1
	If Not rsPrintFormArtCla.Eof Then
		%>
			<form method="post" action="?c=s">
			��<%=rsPrintFormArtCla("Name")%>�� �µ����ƣ�<br />
			<input type="hidden" name="c_id" value="<%=rsPrintFormArtCla("id")%>" />
			<input type="text" name="c_name" value="<%=rsPrintFormArtCla("Name")%>" /><br />
			<input type="submit" value="�޸�" />
			</form>
		<%
	End If
	rsPrintFormArtCla.Close
	Set rsPrintFormArtCla=nothing
End Sub

'������Ƽ����ӱ�
Sub M_PrintFormLink(id)
	sqlCommand="Select * From [Link] Where [id]=" & id
	Set rsPrintFormArtCla=Server.CreateObject("ADODB.RecordSet")
	rsPrintFormArtCla.Open sqlCommand,Conn,1,1
	If Not rsPrintFormArtCla.Eof Then
		%>
			<form method="post" action="?c=s">
			<input type="hidden" name="c_id" value="<%=rsPrintFormArtCla("id")%>" />
			��<%=rsPrintFormArtCla("Name")%>�� �µ����ƣ�<br />
			<input type="text" name="c_name" value="<%=rsPrintFormArtCla("Name")%>" /><br />
			��<%=rsPrintFormArtCla("Name")%>�� �µ�URL��<br />
			<input type="text" name="c_url" value="<%=rsPrintFormArtCla("URL")%>" /><br />
			<input type="submit" value="�޸�" />
			</form>
		<%
	End If
	rsPrintFormArtCla.Close
	Set rsPrintFormArtCla=nothing
End Sub

'������·���ѡ����
Sub M_ArticleClass(cid)
	sqlCommand="Select * From [ArticleClass]"
	Set rsArticleClass=Server.CreateObject("ADODB.RecordSet")
	rsArticleClass.Open sqlCommand,Conn,1,1
	Response.Write("<select name=""classT"">")
	While Not rsArticleClass.Eof
		%><option value="<%=rsArticleClass("id")%>"><%=rsArticleClass("Name")%></option><%
		rsArticleClass.movenext
	WEnd
	Response.Write("</select>")
	rsArticleClass.Close
	Set rsArticleClass = Nothing
End Sub

'����ϵͳ����
Sub SaveInfKey(Key,kvalue)
	sqlCommand="Select * From [Inf] Where [Key]='"&Key&"'"
	Set rsSaveInf=Server.CreateObject("ADODB.RecordSet")
	rsSaveInf.Open sqlCommand,Conn,1,3
	If not rsSaveInf.eof Then
		rsSaveInf("Value")=kvalue & ""
		rsSaveInf.update
	End If
	rsSaveInf.Close
	Set rsSaveInf = Nothing
End Sub

'�޸Ļ�������
Sub SaveInf()
	If Request.Form("Page_Title")<>"" Then Call SaveInfKey("Page_Title",Request.Form("Page_Title"))
	If Request.Form("Page_Keywords")<>"" Then Call SaveInfKey("Page_Keywords",Request.Form("Page_Keywords"))
	If Request.Form("Page_Description")<>"" Then Call SaveInfKey("Page_Description",Request.Form("Page_Description"))
	If Request.Form("Page_Title_1")<>"" Then Call SaveInfKey("Page_Title_1",Request.Form("Page_Title_1"))
	If Request.Form("Page_Title_2")<>"" Then Call SaveInfKey("Page_Title_2",Request.Form("Page_Title_2"))
	If Request.Form("Site_Root")<>"" Then Call SaveInfKey("Site_Root",Request.Form("Site_Root"))
	If Request.Form("Num_DefaultArt")<>"" Then Call SaveInfKey("Num_DefaultArt",Request.Form("Num_DefaultArt"))
	If Request.Form("Button_Comment")<>"" Then Call SaveInfKey("Button_Comment",Request.Form("Button_Comment"))
		If Request.Form("Page_CSS")<>"" Then Call SaveInfKey("Page_CSS",Request.Form("Page_CSS"))
	Call SaveInfKey("Footer_Code",Request.Form("Footer_Code"))
	Call SaveInfKey("Blogger_Name",Request.Form("Blogger_Name"))
	Call SaveInfKey("Blogger_Mail",Request.Form("Blogger_Mail"))
	Call SaveInfKey("Blogger_Con",Request.Form("Blogger_Con"))
	%><p class="meta">���������޸ĳɹ�</p><%
End Sub

'��ȡԶ����ҳ����
function getHTTPPage(url) 
	dim Http 
	set Http=server.createobject("MSXML2.XMLHTTP") 
	Http.open "GET",url,false 
	Http.send() 
	if Http.readystate<>4 then 
		exit function 
	end if 
	getHTTPPage=bytesToBSTR(Http.responseBody,"GB2312")
	set http=nothing
	if err.number<>0 then err.Clear 
end function

'ת������
Function BytesToBstr(body,Cset) 
	dim objstream 
	set objstream = Server.CreateObject("adodb.stream") 
	objstream.Type = 1 
	objstream.Mode =3 
	objstream.Open 
	objstream.Write body 
	objstream.Position = 0 
	objstream.Type = 2 
	objstream.Charset = Cset 
	BytesToBstr = objstream.ReadText 
	objstream.Close 
	set objstream = nothing 
End Function 

'����°汾
function ChkVersion()
    gv=getHTTPPage("http://suomblog.cn/sBlogNewestVersion.asp")
	if gv=GetInf("Sys_V") then
		ChkVersion="v" & GetInf("Sys_V") & " [����]"
	else
		ChkVersion="v" & GetInf("Sys_V") & " <strong><a href=""http://suomblog.cn"" target=""_bank"">[���°汾 v" & gv & "]</a></strong>"
	end if
end function

'�޸����ݿ�
function RDataBase()
	Call NewInf("Page_Title","sBlog",true)
	Call NewInf("Page_Keywords","",true)
	Call NewInf("Page_Description","",true)
	Call NewInf("Page_Title_1","sBlog",true)
	Call NewInf("Page_Title_2","�����õĸ��˲���ϵͳ",true)
	Call NewInf("Site_Root","",true)
	Call NewInf("Num_DefaultArt","7",true)
	Call NewInf("Button_Comment","1",true)
	Call NewInf("Sys_V","1.2",false)
	Call NewInf("Footer_Code","",true)
	Call NewInf("Blogger_Name","�������",true)
	Call NewInf("Blogger_Mail","�������",true)
	Call NewInf("Blogger_Con","�������",true)
	Call NewInf("Page_CSS","css/Default.css",true)
	Call NewInf("Site_MainPage","",true)
	%><p class="meta">���ݿ����޸�������������⣬����ϵ�������ߡ�</p><%
end function

'����µ�Inf��ֵ
function NewInf(f_Key,f_value,f)
	sqlCommand="Select * From [Inf] Where [Key]='" & f_Key & "'"
	Set rsNI=Server.CreateObject("ADODB.RecordSet")
	rsNI.Open sqlCommand,Conn,1,3
	NewInf=false
        If f Then
	    If rsNI.eof Then
		rsNI.Addnew
		rsNI("Key")=f_Key
		rsNI("Value")=f_value
		rsNI.UpDate
		NewInf=Ture
	    End If
        Else
            If Not rsNI.eof Then
		rsNI("Value")=f_value
		rsNI.UpDate
		NewInf=Ture
            End If
        End If
	rsNI.Close
	Set rsNI = Nothing
end function
%>
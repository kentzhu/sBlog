<%
If Session("user_Name")="" Then response.End()
%>

<!--#include FILE="Mag_upfile_class.asp"-->
<%

Server.ScriptTimeOut=9000
dim upfile,formPath,ServerPath,FSPath,formName,FileName,oFile,upfilecount,TmpPath
upfilecount=0

set upfile=new upfile_class 
upfile.NoAllowExt="asa;cdx;cer;hta;htm;html;aspx;exe;cs;vb;js;php;py;jsp;cmd;"	
upfile.GetData (10240000)

%>
<html>
<head>
<title>�ļ��ϴ�</title>
<style type="text/css">
<!--
.p9{ font-size: 9pt; font-family: ���� }
-->
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body leftmargin="20" topmargin="20" class="p9">

<hr size=1 noshadow width=300 align=left>

<%




if upfile.isErr then  '�������
    select case upfile.isErr
	case 1
		Response.Write "��û���ϴ�����?"
	case 2
		Response.Write "���ϴ����ļ��������10M������"
	end select
else

	FSPath=GetFilePath(Server.mappath("Mag_upfile.asp"),"\")'ȡ�õ�ǰ�ļ��ڷ�����·��
	ServerPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/")'ȡ������վ�ϵ�λ��
	
	FSPath=FSPath & "upload/"
	for each formName in upfile.file
		set oFile=upfile.file(formname)
		FileName=upfile.form(formName)
		if not FileName>"" then  FileName=oFile.filename
		FileName=left(right(timer,8),5)& "-" & FileName
		upfile.SaveToFile formname,FSPath&FileName
		
		'�ļ����ƣ�oFile.FilePath&oFile.FileName
		'�ļ���С��oFile.filesize KB 
		
		TmpPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/src/plugins/")
		TmpPath=TmpPath & "upload/" & FileName
		
		if upfile.iserr then 
			Response.Write upfile.errmessage
		else
			upfilecount=upfilecount+1
			Response.Write "�ϴ��ɹ�"
		end if
		set oFile=nothing
	next

end if

If request.QueryString("type")="pho" Then
	Response.write "<script type=text/javascript>parent.KE.plugin[""image""].insert("""&upfile.form("id")&""","""&TmpPath&""","""&upfile.form("imgTitle")&""","""&upfile.form("imgWidth")&""","""&upfile.form("imgHeight")&""","""&upfile.form("imgBorder")&""");</script>"
Else

End If

set upfile=nothing  'ɾ���˶���
Response.End()
%>
</body>
</html>
<%
set upfile=nothing  'ɾ���˶���

Response.End()

function GetFilePath(FullPath,str)
  If FullPath <> "" Then
    GetFilePath = left(FullPath,InStrRev(FullPath, str))
    Else
    GetFilePath = ""
  End If
End function
%>
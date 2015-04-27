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
<title>文件上传</title>
<style type="text/css">
<!--
.p9{ font-size: 9pt; font-family: 宋体 }
-->
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body leftmargin="20" topmargin="20" class="p9">

<hr size=1 noshadow width=300 align=left>

<%




if upfile.isErr then  '如果出错
    select case upfile.isErr
	case 1
		Response.Write "你没有上传数据?"
	case 2
		Response.Write "你上传的文件超出最大10M的限制"
	end select
else

	FSPath=GetFilePath(Server.mappath("Mag_upfile.asp"),"\")'取得当前文件在服务器路径
	ServerPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/")'取得在网站上的位置
	
	FSPath=FSPath & "upload/"
	for each formName in upfile.file
		set oFile=upfile.file(formname)
		FileName=upfile.form(formName)
		if not FileName>"" then  FileName=oFile.filename
		FileName=left(right(timer,8),5)& "-" & FileName
		upfile.SaveToFile formname,FSPath&FileName
		
		'文件名称：oFile.FilePath&oFile.FileName
		'文件大小：oFile.filesize KB 
		
		TmpPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/src/plugins/")
		TmpPath=TmpPath & "upload/" & FileName
		
		if upfile.iserr then 
			Response.Write upfile.errmessage
		else
			upfilecount=upfilecount+1
			Response.Write "上传成功"
		end if
		set oFile=nothing
	next

end if

If request.QueryString("type")="pho" Then
	Response.write "<script type=text/javascript>parent.KE.plugin[""image""].insert("""&upfile.form("id")&""","""&TmpPath&""","""&upfile.form("imgTitle")&""","""&upfile.form("imgWidth")&""","""&upfile.form("imgHeight")&""","""&upfile.form("imgBorder")&""");</script>"
Else

End If

set upfile=nothing  '删除此对象
Response.End()
%>
</body>
</html>
<%
set upfile=nothing  '删除此对象

Response.End()

function GetFilePath(FullPath,str)
  If FullPath <> "" Then
    GetFilePath = left(FullPath,InStrRev(FullPath, str))
    Else
    GetFilePath = ""
  End If
End function
%>
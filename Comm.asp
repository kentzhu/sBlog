

<%
'=================================================
StrDbName="Data/sBlog_Basic.mdb"        '���ݿ������Լ�·��
'=================================================

TimeStar=Timer

'���ӵ����ݿ�
Public Conn
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath(StrDbName)
Conn.Open
If Request.QueryString("logout")="on" Then Call Logout()
If PCHK="ON" Then
	If Session("user_Name")="" Then
		Call LoginChk()
		If Session("user_Name")="" Then
			Session("Error_Login")="��½ʧ��"
			Response.Redirect("Default.asp")
		End If
	End If
End If

CNN_Site_Root=GetInf("Site_Root")
%>
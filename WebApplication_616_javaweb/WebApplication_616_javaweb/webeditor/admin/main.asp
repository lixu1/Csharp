<!--#include file = "private.asp"-->



<%

sPosition = sPosition & "��̨������ҳ"

Call Header()
Call Content()
Call Footer()




Sub Content()
	%>

	<table border=0 cellspacing=1 align=center class=navi>
	<tr><th><%=sPosition%></th></tr>
	</table>

	<br>
	<br>

	<table border=0 cellspacing=1 align=center class=list>
	<tr><th colspan=2>���������йز���</th><th colspan=2>���֧���йز���</th></tr>
	<tr>
	  <td>��ǰ�汾��</td>
	  <td>4.4 asp �汾���޷���������汾��</td>
	  <td>�������ڣ�</td>
	  <td>2007-1-20</td>
	  </tr>
	<tr>
		<td width="15%">����������</td>
		<td width="45%"><%=Request.ServerVariables("SERVER_NAME")%></td>
		<td width="20%">ADO ���ݶ���</td>
		<td width="20%"><%=Get_ObjInfo("adodb.connection", 1)%></td>
	</tr>
	<tr>
		<td width="15%">������IP��</td>
		<td width="45%"><%=Request.ServerVariables("LOCAL_ADDR")%></td>
		<td width="20%">FSO �ı��ļ���д��</td>
		<td width="20%"><%=Get_ObjInfo("scripting.filesystemobject", 0)%></td>
	</tr>
	<tr>
		<td width="15%">�������˿ڣ�</td>
		<td width="45%"><%=Request.ServerVariables("SERVER_PORT")%></td>
		<td width="20%">Stream �ļ�����</td>
		<td width="20%"><%=Get_ObjInfo("Adodb.Stream", 0)%></td>
	</tr>
	<tr>
		<td width="15%">������ʱ�䣺</td>
		<td width="45%"><%=Now()%></td>
		<td width="20%">Jmail �ʼ��շ���</td>
		<td width="20%"><%=Get_ObjInfo("JMail.SMTPMail", 1)%></td>
	</tr>
	<tr>
		<td width="15%">IIS�汾��</td>
		<td width="45%"><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
		<td width="20%">ASPmail ���ţ�</td>
		<td width="20%"><%=Get_ObjInfo("SMTPsvg.Mailer", 1)%></td>
	</tr>
	<tr>
		<td width="15%">����������ϵͳ��</td>
		<td width="45%"><%=Request.ServerVariables("OS")%></td>
		<td width="20%">CDONTS ����SMTP���ţ�</td>
		<td width="20%"><%=Get_ObjInfo("CDONTS.NewMail", 1)%></td>
	</tr>
	<tr>
		<td width="15%">�ű���ʱʱ�䣺</td>
		<td width="45%"><%=Server.ScriptTimeout%> ��</td>
		<td width="20%">LyfUpload �ϴ������</td>
		<td width="20%"><%=Get_ObjInfo("LyfUpload.UploadFile", 1)%></td>
	</tr>
	<tr>
		<td width="15%">վ������·����</td>
		<td width="45%"><%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
		<td width="20%">AspUpload �ϴ������</td>
		<td width="20%"><%=Get_ObjInfo("Persits.Upload.1", 1)%></td>
	</tr>
	<tr>
		<td width="15%">������CPU������</td>
		<td width="45%"><%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</td>
		<td width="20%">SA-FileUp �ϴ������</td>
		<td width="20%"><%=Get_ObjInfo("SoftArtisans.FileUp", 1)%></td>
	</tr>
	<tr>
		<td width="15%">�������������棺</td>
		<td width="45%"><%=ScriptEngine & "/" & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion %></td>
		<td width="20%">AspJpeg ͼ���������</td>
		<td width="20%"><%=Get_ObjInfo("Persits.Jpeg",1)%></td>
	</tr>

	</table>

	
	<%
End Sub

Function Get_ObjInfo(obj, ver)
	On Error Resume Next
	Dim objTest, sTemp
	Set objTest = Server.CreateObject(obj)
	If Err.Number <> 0 Then
		Err.Clear
		Get_ObjInfo = "<font class=red><b>��</b></font>&nbsp;<font class=gray>��֧��</font>"
	Else
		sTemp = ""
		If ver = 1 Then
			sTemp = objTest.version
			If IsNull(sTemp) Then sTemp = objTest.about
			sTemp = Replace(sTemp, "Version", "")
			sTemp = "&nbsp;<font class=tims><font class=blue>" & sTemp & "</font></font>"
		End If
		Get_ObjInfo = "<b>��</b>&nbsp;<font class=gray>֧��</font>" & sTemp
	End If
	Set objTest = Nothing
	If Err.Number <> 0 Then Err.Clear
End Function

%>
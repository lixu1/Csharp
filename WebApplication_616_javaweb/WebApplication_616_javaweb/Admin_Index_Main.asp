<%@language=vbscript codepage=936 %>
<%
'option explicit
'response.buffer=true	
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/check.asp"-->
<%


call CloseConn()

Dim theInstalledObjects(17)
theInstalledObjects(0) = "MSWC.AdRotator"
theInstalledObjects(1) = "MSWC.BrowserType"
theInstalledObjects(2) = "MSWC.NextLink"
theInstalledObjects(3) = "MSWC.Tools"
theInstalledObjects(4) = "MSWC.Status"
theInstalledObjects(5) = "MSWC.Counters"
theInstalledObjects(6) = "IISSample.ContentRotator"
theInstalledObjects(7) = "IISSample.PageCounter"
theInstalledObjects(8) = "MSWC.PermissionChecker"
theInstalledObjects(9) = "Scripting.FileSystemObject"
theInstalledObjects(10) = "adodb.connection"
    
theInstalledObjects(11) = "SoftArtisans.FileUp"
theInstalledObjects(12) = "SoftArtisans.FileManager"
theInstalledObjects(13) = "JMail.SMTPMail"
theInstalledObjects(14) = "CDONTS.NewMail"
theInstalledObjects(15) = "Persits.MailSender"
theInstalledObjects(16) = "LyfUpload.UploadFile"
theInstalledObjects(17) = "Persits.Upload.1"
%>
<html>
<head>
<title>��̨������ҳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Admin_Style.css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
<tr align="center">
    <td height=25 colspan=2 class="topbg"><strong>��̨������ҳ</strong></tr>
<tr> 
    <td width="17%"  class="tdbg" height=23>�����£�2010��6��</td>
    <td width="83%" class="tdbg">�һ�½�����ҵİٶȿռ����淢��һЩʹ�ð����ĵ��������ʲô�õĽ�������Ҳ���Ը������ԡ��ҵİٶȿռ�<a href="http://hi.baidu.com/bangbangnt" target="_blank">bangbangnt</a></td>
  </tr>
  </table>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center"> 
    <td height=25 colspan=2 class="topbg"><strong>�� �� �� �� Ϣ</strong><tr> 
  <tr> 
    <td width="50%"  class="tdbg" height=23>���������ͣ�<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
    <td width="50%" class="tdbg">�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>վ������·����<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
    <td width="50%" class="tdbg">���ݿ��ַ��</td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>FSO�ı���д��
      <%If Not IsObjInstalled(theInstalledObjects(9)) Then%>
      <font color="red"><b>��</b></font>
      <%else%>
      <b>��</b>
      <%end if%></td>
    <td width="50%" class="tdbg">���ݿ�ʹ�ã�
      <%If Not IsObjInstalled(theInstalledObjects(10)) Then%>
      <font color="red"><b>��</b></font>
      <%else%>
      <b>��</b>
      <%end if%></td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>Jmail���֧�֣�
      <%If Not IsObjInstalled(theInstalledObjects(13)) Then%>
      <font color="red"><b>��</b></font>
      <%else%>
      <b>��</b>
      <%end if%></td>
    <td width="50%" class="tdbg">CDONTS���֧�֣�
      <%If Not IsObjInstalled(theInstalledObjects(14)) Then%>
      <font color="red"><b>��</b></font>
      <%else%>
      <b>��</b>
      <%end if%></td>
  </tr>
</table>
<br>
<div align="center">
  Network-based courses Management System<BR>
</div>
</body>
</html>
<%
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
%>
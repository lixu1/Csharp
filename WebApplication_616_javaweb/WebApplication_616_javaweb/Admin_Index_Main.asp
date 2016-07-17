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
<title>后台管理首页</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Admin_Style.css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
<tr align="center">
    <td height=25 colspan=2 class="topbg"><strong>后台管理首页</strong></tr>
<tr> 
    <td width="17%"  class="tdbg" height=23>最后更新：2010年6月</td>
    <td width="83%" class="tdbg">我会陆续在我的百度空间里面发布一些使用帮助文档，大家有什么好的建议和意见也可以给我留言。我的百度空间<a href="http://hi.baidu.com/bangbangnt" target="_blank">bangbangnt</a></td>
  </tr>
  </table>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center"> 
    <td height=25 colspan=2 class="topbg"><strong>服 务 器 信 息</strong><tr> 
  <tr> 
    <td width="50%"  class="tdbg" height=23>服务器类型：<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
    <td width="50%" class="tdbg">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>站点物理路径：<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
    <td width="50%" class="tdbg">数据库地址：</td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>FSO文本读写：
      <%If Not IsObjInstalled(theInstalledObjects(9)) Then%>
      <font color="red"><b>×</b></font>
      <%else%>
      <b>√</b>
      <%end if%></td>
    <td width="50%" class="tdbg">数据库使用：
      <%If Not IsObjInstalled(theInstalledObjects(10)) Then%>
      <font color="red"><b>×</b></font>
      <%else%>
      <b>√</b>
      <%end if%></td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>Jmail组件支持：
      <%If Not IsObjInstalled(theInstalledObjects(13)) Then%>
      <font color="red"><b>×</b></font>
      <%else%>
      <b>√</b>
      <%end if%></td>
    <td width="50%" class="tdbg">CDONTS组件支持：
      <%If Not IsObjInstalled(theInstalledObjects(14)) Then%>
      <font color="red"><b>×</b></font>
      <%else%>
      <b>√</b>
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
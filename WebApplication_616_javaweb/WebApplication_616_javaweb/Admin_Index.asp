<%
if session("AdminName") = "" then
    response.Redirect "Admin_login.asp"
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��������Ŀγ̹���ϵͳ (Network-based courses Management System)</title>
</head>
<frameset id="frame" framespacing="0" border="false" cols="180,*" frameborder="1" scrolling="yes">
	<frame name="left" scrolling="auto" marginwidth="0" marginheight="0" src="Admin_Index_Left.asp">
	<frameset framespacing="0" border="false" rows="35,*" frameborder="0" scrolling="yes">
		<frame name="top" scrolling="no" src="Admin_Index_Top.asp" target="main">
		<frame name="main" scrolling="auto" src="Admin_Index_Main.asp">
	</frameset>
</frameset>
<noframes>
  <body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
  <p>���������汾���ͣ�������ϵͳҪ��IE5�����ϰ汾����ʹ�ñ�ϵͳ��</p>
  </body>
</noframes>
</html>

<%
'--------���岿��------------------
Dim Fy_Post,Fy_Get,Fy_In,Fy_Inf,Fy_Xh,Fy_db,Fy_dbstr
'�Զ�����Ҫ���˵��ִ�,�� "��" �ָ�
Fy_In = "'��;��and��exec��insert��select��delete��update��count��*��%��chr��mid��master��truncate��char��declare"
'----------------------------------
%>

<%
Fy_Inf = split(Fy_In,"��")
'--------POST����------------------
If Request.Form<>"" Then
For Each Fy_Post In Request.Form

For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh))<>0 Then
Response.Write "<Script Language=JavaScript>alert('�벻Ҫ�ڲ����а����Ƿ��ַ���\n\nͬʱ�����Ѿ���¼������IP��');</Script>"
Response.Write "�Ƿ�������ϵͳ�������¼�¼��<br>"
Response.Write "�����ɣУ�"&Request.ServerVariables("REMOTE_ADDR")&"<br>"
Response.Write "����ʱ�䣺"&Now&"<br>"
Response.Write "����ҳ�棺"&Request.ServerVariables("URL")&"<br>"
Response.Write "�ύ��ʽ���Уϣӣ�<br>"
Response.Write "�ύ������"&Fy_Post&"<br>"
Response.Write "�ύ���ݣ�"&Request.Form(Fy_Post)
Response.End
End If
Next

Next
End If
'----------------------------------

'--------GET����-------------------
If Request.QueryString<>"" Then
For Each Fy_Get In Request.QueryString

For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
Response.Write "<Script Language=JavaScript>alert('�벻Ҫ�ڲ����а����Ƿ��ַ���\n\nͬʱ�����Ѿ���¼������IP��');</Script>"
Response.Write "�Ƿ�������ϵͳ�������¼�¼��<br>"
Response.Write "�����ɣУ�"&Request.ServerVariables("REMOTE_ADDR")&"<br>"
Response.Write "����ʱ�䣺"&Now&"<br>"
Response.Write "����ҳ�棺"&Request.ServerVariables("URL")&"<br>"
Response.Write "�ύ��ʽ���ǣţ�<br>"
Response.Write "�ύ������"&Fy_Get&"<br>"
Response.Write "�ύ���ݣ�"&Request.QueryString(Fy_Get)
Response.End
End If
Next
Next
End If
%>
<%Function formatierung(strMessage)
  Dim strTemp
	strTemp = Server.HTMLEncode(strMessage)
	strTemp = Replace(strTemp, "       ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, "      ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, "     ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, "    ", "&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, "   ", "&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, vbCrLf, "<BR>" & vbCrLf, 1, -1, 1)
	formatierung = strTemp
End Function
Function copyrights%>
<!--��Ȩ��Ϣ�������޸�-->
<p style="line-height: 150%" align=center>Java��Ʒ�γ���վ &copy; 2010 All Rights Reserved<br></p>
<%End Function%>
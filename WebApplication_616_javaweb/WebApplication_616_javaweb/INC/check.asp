<%
dim ComeUrl,cUrl
ComeUrl=lcase(trim(request.ServerVariables("HTTP_REFERER")))
action = trim(request("action"))
if action <> "admintest" then
cUrl=lcase(trim("http://" & Request.ServerVariables("SERVER_NAME") & request.ServerVariables("SCRIPT_NAME")))
if ComeUrl="" then
	response.write "<br><p align=center><font color='red'>�Բ���Ϊ��ϵͳ��ȫ��������ֱ�������ַ���ʱ�ϵͳ�ĺ�̨����ҳ�档</font></p>"
	response.end
else
	if lcase(left(ComeUrl,instrrev(ComeUrl,"/")))<>lcase(left(cUrl,instrrev(cUrl,"/"))) and instr(cUrl,"127.0.0.1")=0 then
		response.write "<br><p align=center><font color='red'>�Բ���Ϊ��ϵͳ��ȫ����������ⲿ���ӵ�ַ���ʱ�ϵͳ�ĺ�̨����ҳ�档</font></p>"
		response.end
	end if
end if
end if
%>
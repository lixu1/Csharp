<%
dim SQL_injdata 
SQL_injdata = "'|and|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare|;|@|#|or" 
SQL_inj = split(SQL_injdata,"|") 

If Request.QueryString<>"" Then 
For Each SQL_Get In Request.QueryString 
For SQL_Data=0 To Ubound(SQL_inj) 
if instr(Request.QueryString(SQL_Get),SQL_inj(SQL_Data))>0 Then 
Response.Write "<Script Language=javascript>window.history.back(-1);</Script>" 
Response.end 
end if 
next 
Next 
End If 

'�������Ǿ�ʵ����get�����ע������أ��������ǻ�Ҫ����post�����������ǻ��ü�������request.form,���Ҳ����������ʽ���ڵģ�������ֻ��Ҫ�ٽ�һ��ѭ���жϼ��ɡ��������� 
If Request.Form<>"" Then 
For Each SQL_Post In Request.Form 
For SQL_Data=0 To Ubound(SQL_inj) 
if instr(Request.Form(SQL_Post),SQL_inj(SQL_Data))>0 Then 
Response.Write "<Script Language=javascript>window.history.back(-1);</Script>" 
Response.end 
end if 
next 
next 
end if 
%>
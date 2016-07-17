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

'这样我们就实现了get请求的注入的拦截，但是我们还要过滤post请求，所以我们还得继续考虑request.form,这个也是以数组形式存在的，，我们只需要再进一次循环判断即可。代码如下 
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
<%
'--------定义部份------------------
Dim Fy_Post,Fy_Get,Fy_In,Fy_Inf,Fy_Xh,Fy_db,Fy_dbstr
'自定义需要过滤的字串,用 "枫" 分隔
Fy_In = "'枫;枫and枫exec枫insert枫select枫delete枫update枫count枫*枫%枫chr枫mid枫master枫truncate枫char枫declare"
'----------------------------------
%>

<%
Fy_Inf = split(Fy_In,"枫")
'--------POST部份------------------
If Request.Form<>"" Then
For Each Fy_Post In Request.Form

For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh))<>0 Then
Response.Write "<Script Language=JavaScript>alert('请不要在参数中包含非法字符！\n\n同时我们已经记录下您的IP！');</Script>"
Response.Write "非法操作！系统做了如下记录↓<br>"
Response.Write "操作ＩＰ："&Request.ServerVariables("REMOTE_ADDR")&"<br>"
Response.Write "操作时间："&Now&"<br>"
Response.Write "操作页面："&Request.ServerVariables("URL")&"<br>"
Response.Write "提交方式：ＰＯＳＴ<br>"
Response.Write "提交参数："&Fy_Post&"<br>"
Response.Write "提交数据："&Request.Form(Fy_Post)
Response.End
End If
Next

Next
End If
'----------------------------------

'--------GET部份-------------------
If Request.QueryString<>"" Then
For Each Fy_Get In Request.QueryString

For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
Response.Write "<Script Language=JavaScript>alert('请不要在参数中包含非法字符！\n\n同时我们已经记录下您的IP！');</Script>"
Response.Write "非法操作！系统做了如下记录↓<br>"
Response.Write "操作ＩＰ："&Request.ServerVariables("REMOTE_ADDR")&"<br>"
Response.Write "操作时间："&Now&"<br>"
Response.Write "操作页面："&Request.ServerVariables("URL")&"<br>"
Response.Write "提交方式：ＧＥＴ<br>"
Response.Write "提交参数："&Fy_Get&"<br>"
Response.Write "提交数据："&Request.QueryString(Fy_Get)
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
<!--版权信息，请勿修改-->
<p style="line-height: 150%" align=center>Java精品课程网站 &copy; 2010 All Rights Reserved<br></p>
<%End Function%>
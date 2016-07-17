
<!--#include FILE="upload.inc"-->
<html>
<head>
<title>文件上传</title>
<link rel="stylesheet" href="admin/Style.css" type="text/css">
</head>
<body leftmargin="5" topmargin="3" >

<% 
response.write "<FONT color=red>"&upNum&"</font>"
dim upload,file,formName,formPath,iCount,filename,fileExt
set upload=new upload_5xSoft ''建立上传对象

 formPath="UpTop"
 ''在目录后加(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 

for each formName in upload.file ''列出所有上传了的文件
 set file=upload.file(formName)  ''生成一个文件对象
 if file.filesize<10 then
 	response.write "<font size=2>请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>300000 then
 	response.write "<font size=2>文件大小超过了限制 300K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if fileEXT<>".gif" and fileEXT<>".jpg" and fileEXT<>".swf" then
 	response.write "<font size=2>文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if 

 randomize
 ranNum=int(90000*rnd)+10000
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&fileExt
 
' filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&file.FileName
 
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
  file.SaveAs Server.mappath(FileName)   ''保存文件
  	response.write "<SCRIPT>parent.myform.PicAddress.value='\n"&filename&"'</SCRIPT>"
 end if
 set file=nothing
next
set upload=nothing  ''删除此对象
dim upload_sn
		upload_sn="http://"&LCase(Replace(Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL"),split(request.ServerVariables("SCRIPT_NAME"),"/")(ubound(split(request.ServerVariables("SCRIPT_NAME"),"/"))),""))
call Htmend(upload_sn & FileName)

sub HtmEnd(Msg)
 set upload=nothing
 'response.write ("<table width=100% ><tr><td height=40 align=center ><a href=# onclick=history.go(-1)>继续上传</a></td></tr></table>")
 'response.end
end sub
%>
</body>
</html>
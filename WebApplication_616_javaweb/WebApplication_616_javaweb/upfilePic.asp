
<!--#include FILE="upload.inc"-->
<html>
<head>
<title>�ļ��ϴ�</title>
<link rel="stylesheet" href="admin/Style.css" type="text/css">
</head>
<body leftmargin="5" topmargin="3" >

<% 
response.write "<FONT color=red>"&upNum&"</font>"
dim upload,file,formName,formPath,iCount,filename,fileExt
set upload=new upload_5xSoft ''�����ϴ�����

 formPath="UpTop"
 ''��Ŀ¼���(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 

for each formName in upload.file ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
 if file.filesize<10 then
 	response.write "<font size=2>����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>300000 then
 	response.write "<font size=2>�ļ���С���������� 300K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if fileEXT<>".gif" and fileEXT<>".jpg" and fileEXT<>".swf" then
 	response.write "<font size=2>�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if 

 randomize
 ranNum=int(90000*rnd)+10000
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&fileExt
 
' filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&file.FileName
 
 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
  file.SaveAs Server.mappath(FileName)   ''�����ļ�
  	response.write "<SCRIPT>parent.myform.PicAddress.value='\n"&filename&"'</SCRIPT>"
 end if
 set file=nothing
next
set upload=nothing  ''ɾ���˶���
dim upload_sn
		upload_sn="http://"&LCase(Replace(Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL"),split(request.ServerVariables("SCRIPT_NAME"),"/")(ubound(split(request.ServerVariables("SCRIPT_NAME"),"/"))),""))
call Htmend(upload_sn & FileName)

sub HtmEnd(Msg)
 set upload=nothing
 'response.write ("<table width=100% ><tr><td height=40 align=center ><a href=# onclick=history.go(-1)>�����ϴ�</a></td></tr></table>")
 'response.end
end sub
%>
</body>
</html>
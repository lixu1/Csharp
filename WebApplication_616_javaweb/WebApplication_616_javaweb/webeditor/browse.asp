<%@ language=VBScript CODEPAGE=936%>
<%
'Option Explicit
response.buffer= True
%>
<html>
<head>
<title>文件管理-设计在线4.8专用版本</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<base target='_self'>
<style>
A:link {

	COLOR: #000000; TEXT-DECORATION: none;/*FONT-SIZE: 12px; */
}
A:visited {
	 COLOR: #000000; TEXT-DECORATION: none;/*FONT-SIZE: 12px;*/
}
A:active {
	COLOR: #000000; TEXT-DECORATION: none;/*FONT-SIZE: 12px; */
}
A:hover {
	 COLOR: #000000; TEXT-DECORATION: underline;/*FONT-SIZE: 12px;*/
}
.mydiv{float:left; width:170;height:176;text-align:center;overflow:hidden;
padding:5px 5px 5px 5px;}
.myli{float:left; width:160;height:150;text-align:center;overflow:hidden;
background:#FAFBF0;}
.myspan{float:left; width:160;height:20;text-align:center;overflow:hidden;}
</style>
</head>
<script language="JavaScript">
function DrawImage(ImgD){
   var image=new Image();
   image.src=ImgD.src;
   if(image.width>0 && image.height>0){
    flag=true;
    if(image.width/image.height>= 160/160){
     if(image.width>160){  
     ImgD.width=160;
     ImgD.height=(image.height*160)/image.width;
     }else{
     ImgD.width=image.width;  
     ImgD.height=image.height;
     }

     }
    else{
     if(image.height>160){  
     ImgD.height=160;
     ImgD.width=(image.width*160)/image.height;     
     }else{
     ImgD.width=image.width;  
     ImgD.height=image.height;
     }

     }
    }
   }
</script>
<body leftmargin="0" topmargin="0">

<%

'Response.Write("the folder in this place:")
'response.Write("<br />")

set fso=server.createobject("Scripting.FileSystemObject") 
if Trim(Request.QueryString("folder"))="" then 
'strpath="."
 strpath="/webeditor/uploadfile/"
else 
strpath=Trim(Request.QueryString("folder"))&"/"
end if

call getfolders(strpath) 
function getfolders(strpath) 
set objfolder=fso.getfolder(server.mappath(strpath)) 
pathul=GetUpDir()

Response.Write("<a href="""&"?folder="&pathul&" "">")

response.Write "<img src=""sysimage/file/parentfolder.gif"" border=0 width=""16"" height=""16"" align=""absmiddle"">"&pathul&   "..</a>" 
Response.Write("<br />")


for each objsubfolder in objfolder.subfolders  
Response.Write("<a href=""?folder="&replace(strpath,".","") & "" &objsubfolder.name&" "">")
Response.write "<img src=""sysimage/file/openfolder.gif"" border=0 width=""16"" height=""16"" align=""absmiddle"">"&replace(strpath,".","") &  objsubfolder.name & "</a>" 
Response.Write("&nbsp;")
'call getfolders (strpath & "\" & objsubfolder.name) 
next 
End function 

'Response.Write("this is write by educms.cn ")
call getfiles(strpath)
'response.Write(server.MapPath("/UploadFiles"))
function getfiles(path)
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    'on error resume next 
    Set objFolder=fso.GetFolder(server.MapPath(path))
    Set objFiles=objFolder.Files
       FolderNuns=objFolder.Files.count
	  ' Response.Write("there have :")
	  ' response.Write(FolderNuns)
	  ' Response.Write("Files<br />there folder: <br />")
	   'response.Write(objFolder.name)
	   Response.Write(" <br />")

    For each objFile in objFiles
	    
        filelist = objFile.name
		'folder=objFile.subfolders
		if LCase(right(filelist,3))="jpg" or LCase(right(filelist,3))="gif" or LCase(right(filelist,3))="bmp" or LCase(right(filelist,3))="png" then
		response.Write("<div class=""mydiv"">")
		response.Write "<span class=""myli"">"
		response.Write "<a href=""#"" onClick=""window.returnValue='"&replace(path,".","")&""&server.HTMLEncode(filelist)&"';window.close();""><img  src="&replace(path,".","")&filelist&" onload=""javascript:DrawImage(this)"" border=0 /></a>"
		response.Write("</span>")
		response.Write("<span class=""myspan"">"&filelist&"</span>") 
		response.Write("</div>")
		end if
    Next
       
    Set objFolder=nothing
    Set fso=nothing
end function

Function GetUpDir()
	Dim strDir,strDir2,i
	strDir=""
	ThisDir=replace(strpath,".","")
	If ThisDir = "" then Exit Function
	strDir2=ThisDir
	strDir2=Split(strDir2,"/")
	for i=0 to Ubound(strDir2)-1
	    if i<Ubound(strDir2)-2 then strDir=strDir & strDir2(i) &"/"
		if i=Ubound(strDir2)-2 then strDir=strDir & strDir2(i)
	next
	GetUpDir=strDir
	'Response.Write("<br />the path is:")
	'response.Write(GetUpDir)
End Function
 %>


<!--#include file="inc/conn.asp"-->
<!--#include file="inc/check.asp"-->
<%
dim ArticleID,Action,sqlDel,rsDel,FoundErr,ErrMsg
ArticleID=trim(request("ArticleID"))
Action=Trim(Request("Action"))
FoundErr=False
if Action="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
end if
if FoundErr=False then
	if Action="Del" then
		call DelArticle()
	end if
end if
if FoundErr=False then
	call CloseConn()
	response.Redirect "Admin_ArticleManage.asp"
else
	call WriteErrMsg()
	call CloseConn()
end if

sub DelArticle()
	if ArticleID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请先选定信息！</li>"
		exit sub
	end if

	if instr(ArticleID,",")>0 then
		ArticleID=replace(ArticleID," ","")
		sqlDel="select * from Article where ArticleID in (" & ArticleID & ")"
	else
		ArticleID=Clng(ArticleID)
		sqlDel="select * from article where ArticleID=" & ArticleID
	end if
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	do while not rsDel.eof
		rsDel.delete
		rsDel.movenext
	loop
	rsDel.close
	set rsDel=nothing
end sub
%>

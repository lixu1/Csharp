<!--#include file="inc/conn.asp"-->
<!--#include file="inc/check.asp"-->
<%
dim ID,Action,sqlDel,rsDel,FoundErr,ErrMsg
ID=trim(request("newsID"))
Action=Trim(Request("Action"))
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
	response.Redirect "Admin_NewsManage.asp"
else
	call WriteErrMsg()
	call CloseConn()
end if

sub DelArticle()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请先选定新闻id！</li>"
		exit sub
	end if

	if instr(ID,",")>0 then
		ID=replace(ID," ","")
		sqlDel="select * from news where newsID in (" & ID & ")"
	else
		ID=Clng(ID)
		sqlDel="select * from news where newsID=" & ID
	end if
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	rsDel.open sqlDel,conn,1,3
	do while not rsDel.eof
		rsDel.delete
		rsDel.update
		rsDel.movenext
	loop
	rsDel.close
	set rsDel=nothing
end sub


%>

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="keywords" content="<%=Sitekeyword%>">
<meta name="description" content="<%=sitename%>">
<META NAME="robots" CONTENT="index, follow">
<META NAME="GOOGLEBOT" CONTENT="INDEX, FOLLOW">
<META NAME="baiduspider" CONTENT="INDEX, FOLLOW">
<META NAME="slurp" CONTENT="INDEX, FOLLOW">
<meta name="date" content="<%=now()%>" scheme="yyyy/mm/dd hh:mm:ss" />
<link href="images/style.css" rel="stylesheet" type="text/css" />

<title><%=SiteTitle%></title>
</head>
<body>
<!--#include file="head.asp"-->
<div class="path">
	<p>����λ�ã�<a href="default.asp"><%=SiteName%></a> > ��վ��ҳ</p>
</div>
<div class="basic btm10">
	<div class="left">
		<div class="living">
			<h2>���¹���</h2>
			<p id="srollingTips">
<%
	sql="select top 5 NewsID,Title,Adddatetime from news where ClassID=3 order by adddate desc"
	set rs=conn.execute(sql)
	do until rs.eof
	NewsID	=	rs("NewsID")
	Title		=	trim(rs("Title"))
	Adddatetime	=	formatdatetime(rs("Adddatetime"),2)
%>      
				<span><a href="news.asp?newsid=<%=newsid%>" target="_blank"><%=title%></a>&nbsp;&nbsp;����ʱ�䣺<em><%=Adddatetime%></em></span>
<%
	rs.movenext
	loop
	rs.close
	set rs=nothing
%>
			</p>
		</div>

		<div class="zh_recommed" id="zh_recommed">
			<h2><span>�γ̼��</span></h2>
		<div class="info">
    <%
		sql="select content from config where flag=1"
		set rs=conn.execute(sql)
		if not rs.eof then
			response.Write(rs("content"))
		end if
		rs.close
		set rs=nothing
		%>
    </div>
		</div>
 
		<div class="zh_news">
			<h2><span>�γ̶�̬</span></h2><a href="newlist.asp?classid=1" target="_blank" class="more">����>></a>
			<ul>
				<%=shownewlist(8,1)%>
			</ul>	
		</div>
		
		<div class="zh_news">
			<h2><span>ҵ����Ѷ</span></h2><a href="newlist.asp?classid=2" target="_blank" class="more">����>></a>
			<ul>
				<%=shownewlist(8,2)%>
			</ul>	
		</div>

		<div class="srolling">
			<h2><span>����ѧ</span></h2>
			<div class="photo_list" id="photos">
				<ul id="showArea">
<%
	sql="select imgname,imgpath,imglink from img where show=true and Role=1 order by sort asc"
	set rs=conn.execute(sql)
	do until rs.eof 
	imgname	=	trim(rs("imgname"))
	imgpath	=	trim(rs("imgpath"))
	imglink	=	trim(rs("imglink"))
	if imglink="" then imglink=imgpath
%>        
					<li><a href="<%=imglink%>" target="_blank"><img src="<%=imgpath%>" alt="<%=imgname%>" /></a><a href="<%=imglink%>" target="_blank"><%=imgname%></a></li>
<%
	rs.movenext
	loop
	rs.close
	set rs=nothing
%>
				</ul>
			</div>
			<span class="prev" id="goleft">�������</span> 
			<span class="prev next" id="goright">���ҹ���</span>
		</div>

		<div class="recomed">
			<h2><span>ѧϰ����</span></h2><a href="newlist.asp?classid=4" target="_blank" class="more">����>></a>
			<ul>
				<%=shownewlist(8,4)%>
			</ul>	
		</div>

		<div class="recomed dismost">
			<h2><span>��������</span></h2><a href="newlist.asp?classid=5" target="_blank" class="more">����>></a>
			<ul>
				<%=shownewlist(8,5)%>
			</ul>	
		</div>
	</div>
	
	<div class="right">
	<!--#include file="right.asp"-->
		<div class="report">
			<ul>
				<li><h2><a href="info.asp?articleid=30">�� ѧ �� ��</a></h2></li>
				<li><h2><a href="info.asp?articleid=31">�� ѧ �� ��</a></h2></li>
				<li><h2><a href="info.asp?articleid=17">�� �� �� ��</a></h2></li>
				<li><h2><a href="info.asp?articleid=26">ʵ ѵ �� ��</a></h2></li>
				<li><h2><a href="info.asp?articleid=20">�� Ԫ ʵ ��</a></h2></li>
				<li><h2><a href="info.asp?articleid=27">�� �� ʵ ѵ</a></h2></li>
			</ul>	
		</div>
<%
Set conn1 = Server.CreateObject("ADODB.Connection")
connstr1="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("bbs/data.asp")
conn1.Open connstr1
sql="select top 10  message,connected,datum_updated from postings where connected<>0 order by datum_updated desc"
	set rs=conn1.execute(sql)
	if not rs.eof then
%>    
		<div class="dislist">
			<h2><span>��������</span></h2><a href="bbs/index.asp" target="_blank" class="more">����>></a>
			<ul>
<%
		do while not rs.eof 
		message		= rs("message")
		connected	=	rs("connected")
		datum_updated	=	formatdatetime(rs("datum_updated"),2)
%>      
				<li><a href="bbs/view.asp?id=<%=connected%>" target="_blank"><%=gotTopic(message,18)%></a><small><%=datum_updated%></small></li>
<%
		rs.movenext
		loop
%>        
			</ul>	
		</div>
<%
	end if
	rs.close
	set rs=nothing
conn1.close
set conn1=nothing
%>
	</div>
	<div class="clear"></div>
</div>
<!--#include file="copyright.asp"-->
<script type="text/javascript" src="jquery.js"></script>
<script type="text/javascript" src="js.js"></script>
</body>
</html>
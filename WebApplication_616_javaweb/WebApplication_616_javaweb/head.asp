<!--#include file="inc/sqlcheck.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<div id="banner">
  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="960" height="130">
    <param name="movie" value="images/top.swf" />
    <param name="quality" value="high" />
    <param name="wmode" value="transparent" />
    <embed src="images/top.swf" quality="high" wmode="opaque" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="960" height="130"></embed>
  </object>
</div>
<div id="navigat">
  <ul id="nav">
  	<li><a href="default.asp">网站首页</a></li>
<%
	sql="select BigID,BigName from bigclass order by Adddate asc"
	set rs=conn.execute(sql)
	do until rs.eof
		BigID		=	trim(rs("BigID"))
		BigName	=	trim(rs("BigName"))
%>	<li><a href="#"><%=BigName%></a>
<%
		sql1="select ArticleID,Title from article where BigID="&BigID&" order by AddDate asc"
		set rs1=conn.execute(sql1)
		if not rs1.eof then
			response.Write("	<ul>")		
			do until rs1.eof
				ArticleID		=	trim(rs1("ArticleID"))
				Title		=	trim(rs1("Title"))
%> 	<li><a href="info.asp?articleid=<%=ArticleID%>"><%=title%></a></li>
<%
			rs1.MoveNext
			loop
			response.Write("	</ul>")
		end if
		rs1.close
		set rs1=nothing
%>
	</li>
<%
	rs.MoveNext
	loop
	rs.close
	set rs=nothing
%>
    <li><a href="#">网上学习</a>
    	<ul>
      <li><a href="newlist.asp?classid=4">学习资料</a></li>
      <li><a href="bbs/index.asp" target="_blank">交流讨论</a></li>
      <li><a href="newlist.asp?classid=5">常见问题</a></li>
      <li><a href="#">单元自测</a></li>
      <li><a href="#">综合模拟</a></li>
      </ul>
    </li>
  	<li><a href="infoshow.asp?flag=3">政策支持</a></li>
  </ul>
</div>
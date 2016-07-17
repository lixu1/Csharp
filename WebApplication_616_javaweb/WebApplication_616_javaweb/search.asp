<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SiteTitle%></title>
<meta name="keywords" content="<%=Sitekeyword%>">
<meta name="description" content="<%=sitename%>">
<META NAME="robots" CONTENT="index, follow">
<META NAME="GOOGLEBOT" CONTENT="INDEX, FOLLOW">
<meta name="searchtitle" content="<%=SiteKeyword%>">
<meta NAME="author" CONTENT="<%=sitename%>">
<meta NAME="copyright" content="<%=sitename%>">
<link href="css/STYLE.CSS" rel="stylesheet" type="text/css" />
</head>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<%
	classID=request.QueryString("ClassID")
	if classID="" then
		className="新闻中心"
	else
		className=GetNewsClass(ClassID)
	end if
%>
<body>
<table width="985" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#FFFFFF"><!--#include file="top.asp" --></td>
  </tr>
  <tr>
    <td align="left" valign="middle" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="195" align="left" valign="top" bgcolor="#FCFCFC"><table width="195" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><img src="pic/index-18.gif" width="195" height="29" /></td>
          </tr>
          <tr>
            <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
                <tr>
                  <td height="100" align="left" valign="top"><%
            set Rs=server.createobject("adodb.recordset")
            sqlStr="select * from newsclass where classID order by classID"
            Rs.open sqlStr,conn,1,1
			if not Rs.eof then
			response.write "<ul class='nav_left'>"
            do while not Rs.eof
            	response.write "<li><a href='news.asp?classID="&Rs("ClassID")&"'><font color='#ED4308'>·</font>"&Rs("ClassName")&"</a></li>"
            	Rs.movenext
            loop
			response.write "</ul>"
			end if
            Rs.close
            %>
				  </td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td><img src="pic/index-2.gif" width="195" height="41" /></td>
          </tr>
          <tr>
            <td height="40" align="center"><img src="pic/index-4.gif" width="187" height="36" /></td>
          </tr>
          <tr>
            <td height="40" align="center"><img src="pic/index-5.gif" width="187" height="36" /></td>
          </tr>
          <tr>
            <td height="40" align="center"><img src="pic/index-6.gif" width="187" height="36" /></td>
          </tr>
          <tr>
            <td height="40" align="center"><img src="pic/index-7.gif" width="187" height="36" /></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table></td>
        <td width="595" align="center" valign="top" class="td_index_1"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="28" background="pic/index-21.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="8%"><img src="pic/index-20.gif" width="45" height="29" /></td>
                <td width="33%" align="left" valign="middle"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="22" valign="bottom"><span class="td_index_4"><%=ClassName%></span></td>
                  </tr>
                </table></td>
                 <td width="59%" align="right">您现在的位置&gt;&gt;<a href=index.asp>首页</a>&gt;&gt;<%if className<>"新闻中心" then response.Write("新闻中心&gt;&gt;")%><%=ClassName%>&nbsp;</td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td align="left"><img src="pic/index-19.gif" width="183" height="13" /></td>
          </tr>
          <tr>
            <td height="10" align="left">&nbsp;</td>
          </tr>
          <tr>
            <td height="500" align="center" valign="top">
		<%
		keyword=request.Form("keyword")
		if keyword="" then Call ShowErr("请输入关键字！","index.asp")
		set Rs=server.CreateObject("adodb.recordset")
		sqlStr="select hotpic,newsid,title,adddatetime from news where title like '%"&keyword&"%' or content like '%"&keyword&"%' order by adddate desc, newsid desc"
		Rs.open sqlStr,conn,1,1
		Rs.pagesize=20
		page =clng(Request.QueryString("page"))
		if page<1 then page=1
		if page>Rs.pagecount then page=Rs.pagecount
		if not Rs.eof then
		rs.AbsolutePage = page
		for i=1 to Rs.pagesize
		%>
			<table width="96%" height="24" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td height="20" align="left" background="pic/index-bg.gif"><font color="#ED4308">·<%if rs("hotpic")<>"" then response.Write("[图]")%></font><a href="news_show.asp?id=<%=Rs("newsid")%>" target="_blank"><%=Rs("title")%></a>&nbsp;<font color="#ED4308">[<%=formatdatetime(Rs("adddatetime"),2)%>]</font></a><%
					adddate=formatdatetime(rs("adddatetime"),2)
					nowdate=formatdatetime(now(),2)
					if adddate=nowdate then response.Write("&nbsp;<img src=pic/icon_new.gif />")
					%></td>
              </tr>
            </table>
	<%
			rs.movenext
			if rs.eof then exit for
		next
		else
			response.Write("没有新闻")
		end if
	%>
			</td>
          </tr>
          <tr>
            <td height="35" align="center" valign="middle" bgcolor="#FCFCFC" class="td_index_5"><%
			if Rs.RecordCount<>0 then
				if rs.pagecount=1 then
					response.Write "上一页&nbsp;&nbsp;"
					response.Write "下一页&nbsp;&nbsp;"
				else
					if page=1 then
						response.Write "上一页&nbsp;&nbsp;"
						response.Write "<a href=?page="& (page+1) &">下一页</a>&nbsp;&nbsp;"
					elseif page=Rs.pagecount then
						response.Write "<a href=?page=" & (page-1) &">上一页</a>&nbsp;&nbsp;"
						response.Write "下一页&nbsp;&nbsp;"
					else
						response.Write "<a href=?page=" & (page-1) &">上一页/a>&nbsp;&nbsp;"
						response.Write "<a href=?page="& (page+1) &">下一页</a>&nbsp;&nbsp;"
					end if
				end if
			end if
			%>
&nbsp;&nbsp;当前页<%=page%>/<%=rs.pagecount%>，&nbsp;共<%=rs.RecordCount%>条信息 &nbsp;20条每页
<%Rs.close%></td>
          </tr>
        </table></td>
        <td width="195" align="left" valign="top"><!--#include file="right.asp"--></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><!--#include file="bottom.asp" --></td>
  </tr>
</table>
</body>
</html>

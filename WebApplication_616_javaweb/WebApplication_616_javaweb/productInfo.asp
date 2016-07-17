<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
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
<link href="css/css.CSS" rel="stylesheet" type="text/css" />
</head>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<body>
<script language=javascript id=clientEventHandlersJS>
<!--
var number=8;

function LMYC() {
var lbmc;
//var treePic;
    for (i=1;i<=number;i++) {
        lbmc = eval('LM' + i);
        //treePic = eval('treePic'+i);
        //treePic.src = 'pic/file.gif';
        lbmc.style.display = 'none';
    }
}
 
function ShowFLT(i) {
    lbmc = eval('LM' + i);
    //treePic = eval('treePic' + i)
    if (lbmc.style.display == 'none') {
        LMYC();
        //treePic.src = 'pic/nofile.gif';
        lbmc.style.display = '';
    }
    else {
        //treePic.src = 'pic/file.gif';
        lbmc.style.display = 'none';
    }
}
//-->
      </script>
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><!--#include file="top.asp" --></td>
  </tr>
  <tr>
    <td align="left" valign="top"><table width="1003" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="287" align="left" valign="top" bgcolor="#F6F6F6" class="index_td1"><!--#include file="left.asp"--></td>
        <td align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="37" background="pic/index_3.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
			<%
			dim j,bigid
			bigid=request.QueryString("bigid")
			smallID=request.QueryString("SmallID")
			j=0
			set Rs=server.CreateObject("adodb.recordset")
			sqlStr="select * from Article where ArticleID>0"
			if clng(BigID)>0 then
				sqlStr=sqlStr & " and bigid=" & clng(bigID)
			end if
			if Clng(SmallID)>0 then
				sqlStr=sqlStr & " and SmallID=" & Clng(SmallID)
			end if
			sqlStr=sqlStr & " order by adddate desc, Ontop,Elite,Articleid desc"
			Rs.open sqlStr,conn,1,1
					if not rs.eof then
						  	set rs1=server.CreateObject("adodb.recordset")
							if bigid>0 then
							sql1="select * from BigClass where BigID="&rs("BigID")
							rs1.open sql1,conn,1,1
							BigName=rs1("BigName")
							rs1.close
							end if
							if smallid>0 then
							sql1="select * from smallclass where SmallID="&rs("smallID")
							rs1.open sql1,conn,1,1
							SmallName=rs1("smallname")
							rs1.close
							end if
							set rs1=nothing
						end if
%>
                <td width="4%"><img src="pic/index_8.gif" width="24" height="37" /></td>
                <td width="22%" class="index_td9">产品信息</td>
                <td width="74%" align="right" style="padding-right:8px; padding-top:7px">您现在的位置：<a href="index.asp">首页</a>&gt;&gt;产品信息
				<%
					if len(bigname)>0 then
								if len(smallname)>0 then
								%>
                        <a href=productinfo.asp?bigID=<%=rs("bigid")%>><font color="#333">>><%=bigname%></font></a><a href=productinfo.asp?smallid=<%=rs("smallid")%>><font color="#333">>><%=smallname%></font></a>
                        <%else%>
                        <a href=productinfo.asp?bigID=<%=rs("bigid")%>><font color="#333">>><%=bigname%></font></a>
                      <%
							end if
							end if
							%></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td class="index_5"><table width="100%" border="0" cellspacing="8" cellpadding="8">
              <tr>
                <td align="center" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
<%
			Rs.pagesize=12
										  
			page =clng(Request.QueryString("page"))
			if page<1 then page=1
			if page>Rs.pagecount then page=Rs.pagecount
										  
			if not Rs.eof then   
			rs.AbsolutePage = page		
			for i=1 to Rs.pagesize 
			%>
                      <td width="20%" align="center" valign="top"><table border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="183" height="133" align="center" style="padding:3px;"><a href="product_show.asp?id=<%=rs("articleID")%>" title="<%=rs("Title")%>"><img src="<%=rs("defaultpic")%>" width="180" height="120" border="0" style="border:1px solid #eee;" /></a></td>
                          </tr>
                          <tr>
                            <td align="center" style="padding:3px;"><a href="product_show.asp?id=<%=rs("articleID")%>" title="<%=rs("Title")%>">
                              <%
					title=rs("Title")
					if len(title)>9 then response.Write(left(title,9)&"…") else response.Write(title)
					%>
                            </a></td>
                          </tr>
                        </table>
                          <%
				if i mod 3=0 then response.Write("</tr><tr>")
				rs.movenext
				if rs.eof then exit for
			next
			else
				response.Write("还没有添加信息!")
			end if
			%>
                      </td>
                    </tr>
                  </table>
                    <span style="padding:10px;">
                    <%
				dim filename
				filename=""
				if clng(bigID)>0 then
					filename=filename & "&BigID=" & bigID&"&smallID="&smallID
				end if
				if Rs.RecordCount<>0 then
					if rs.pagecount=1 then
						response.Write "上一页&nbsp;&nbsp;"
						response.Write "下一页&nbsp;&nbsp;"
					else
						if page=1 then
							response.Write "上一页&nbsp;&nbsp;"
							response.Write "<a href=?page="& (page+1) & filename &">下一页</a>&nbsp;&nbsp;"
						elseif page=Rs.pagecount then
							response.Write "<a href=?page=" & (page-1) & filename &">上一页</a>&nbsp;&nbsp;"
							response.Write "下一页&nbsp;&nbsp;"
						else
							response.Write "<a href=?page=" & (page-1) &filename&">上一页</a>&nbsp;&nbsp;"
							response.Write "<a href=?page="& (page+1) & filename&">下一页</a>&nbsp;&nbsp;"
						end if
					end if
				
				%>
                      &nbsp;&nbsp;当前页<%=page%>/<%=rs.pagecount%>，&nbsp;共<%=rs.RecordCount%>条信息 &nbsp;<%=rs.pagesize%>条信息每页
                      <%
				end if
				Rs.close%>
                  </span></td>
              </tr>
            </table></td>
          </tr>
          
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><!--#include file="foot.asp" --></td>
  </tr>
</table>
</body>
</html>

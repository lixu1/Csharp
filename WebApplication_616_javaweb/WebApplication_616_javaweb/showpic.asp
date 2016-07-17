				<%
					set rs=server.CreateObject("adodb.recordset")
					sql="select top 6 * from article where defaultpic<>'' and bigID=4 order by adddate desc,articleID desc"
					rs.open sql,conn,1,1
					if rs.eof then
						response.Write("还没有信息")
					else
				%>
				<div id="mq" style="width:100%;height:123px;overflow:hidden" onmouseover="iScrollAmount=0"
onmouseout="iScrollAmount=1">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
				<%
						i=0
						do while not rs.eof
					%>
                          <td align="center">
						  <div>
						  <a href="product_show.asp?id=<%=rs("articleID")%>" target="_blank" title="<%=rs("title")%>"><img src="<%=rs("defaultpic")%>" width="71" height="96" border="0" style="border:1px solid #f3f3f3;" /></a></div>
						  <div style="height:25px;">
                          <a href="product_show.asp?id=<%=rs("articleID")%>" target="_blank" title="<%=rs("title")%>"><%
								  title=rs("title")
								  if len(title)>5 then response.Write(left(title,5)) else response.Write(title)
								  %></a></div></td>
						<%
						i=i+1
						if i mod 3 =0 then
							response.Write("</tr><tr>")
						end if
						rs.movenext
						loop
					%>
                        </tr>
                      </table>
</DIV>
		  <%
		  end if
		  rs.close
		  %>
<script>
var oMarquee = document.getElementById("mq"); //滚动对象
var iLineHeight = 123; //单行高度，像素
var iLineCount = 2; //实际行数
var iScrollAmount = 1; //每次滚动高度，像素
function run() {
oMarquee.scrollTop += iScrollAmount;
if ( oMarquee.scrollTop == iLineCount * iLineHeight )
oMarquee.scrollTop = 0;
if ( oMarquee.scrollTop % iLineHeight == 0 ) {
window.setTimeout( "run()", 2000 );
} else {
window.setTimeout( "run()", 50 );
}
}
oMarquee.innerHTML += oMarquee.innerHTML;
window.setTimeout( "run()", 2000 );
</script>
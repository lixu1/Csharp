<div class="basic copy">
	<p class="cmenu"><a href="infoshow.asp?flag=1">课程简介</a>   |  <a href="info.asp?articleid=3">课程负责人</a>  |  <a href="info.asp?articleid=5">教学团队</a>  |  <a href="info.asp?articleid=14">教学获奖</a>  |  <a href="info.asp?articleid=15">课改历程</a> </p>
<%
sql="select Content from config where flag=8"
set rs=conn.execute(sql)
if not rs.eof then
 response.Write(rs("Content"))
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
技术支持：<a href="http://hi.baidu.com/bangbangnt" target="_blank">bangbangnt</a>  <script src="http://s3.cnzz.com/stat.php?id=1978483&web_id=1978483&show=pic" language="JavaScript"></script>
<div style="display:none">http://www.haozhanhui.com</div>
</div>
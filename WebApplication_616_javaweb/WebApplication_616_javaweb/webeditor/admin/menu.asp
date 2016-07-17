<!--#include file = "private.asp"-->


<html>
<head>
<title>eWebEditor</title>
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<link type=text/css href='private.css' rel=stylesheet>
<base target=main>
</head>
<script language=javascript>
<!--
function menu_tree(meval)
{
  var left_n=eval(meval);
  if (left_n.style.display=="none")
  { eval(meval+".style.display='';"); }
  else
  { eval(meval+".style.display='none';"); }
}
-->
</script>
<body>
<center>

  <table cellspacing=0  class="Menu">
  <tr><th align=center  onclick="javascript:menu_tree('left_1');" >≡ 首选服务 ≡</th></tr>
  <tr id='left_1'><td >
    <table width='100%'>
    <tr><td><img border=0 src='images/menu.gif' align=absmiddle>&nbsp;<a href='style.asp'>样式管理</a></td></tr>
    </table>
  </td></tr>
  <tr id='left_1'><td >
    <table width='100%'>
    <tr><td><img border=0 src='images/menu.gif' align=absmiddle>&nbsp;<a href='databak.asp'>数据备份</a></td></tr>
    </table>
  </td></tr>
  </table>

  <table width='90%' height=2><tr ><td></td></tr></table>
  <table cellspacing=0  class="Menu">
  <tr><th align=center  onclick="javascript:menu_tree('left_2');" >≡ 辅助服务 ≡</th></tr>
  <tr id='left_2'><td>
    <table width='100%'>
    <tr><td><img border=0 src='images/menu.gif' align=absmiddle>&nbsp;<a href='main.asp'>后台首页</a></td></tr>
    <tr><td><img border=0 src='images/menu.gif' align=absmiddle>&nbsp;<a onClick="return confirm('提示：您确定要退出系统吗？')" href='login.asp?action=out' target='_parent'>退出后台</a></td></tr>
    <tr>
      <td><img border=0 src='images/menu.gif' align=absmiddle>&nbsp;<a href="modipwd.asp" >修改密码</a></td>
    </tr>
    </table>
  </td></tr>
  </table>
  
  <table width='90%' height=2><tr ><td></td></tr></table>
  <table cellspacing=0  class="Menu">
  <tr><th align=center  >〓 版本信息 〓</th></tr>
  <tr>
    <td align=center><p>eWebEditor<br> 
      4.4（asp版本）</p>
      </td>
  </tr>
  <tr><td align=center></td>
  </tr>
  </table>

</center>
</body>
</html>
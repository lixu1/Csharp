<%connstr1=""
		db="bottom.asa"
		connstr1="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
		Set conn1 = Server.CreateObject("ADODB.Connection")
	conn1.Open connstr1
	If Err Then
		err.Clear
		Set Conn1 = Nothing
		Response.Write "Sorry! 数据库连接出错，请检查连接字串。"
		Response.End
	End If
%>
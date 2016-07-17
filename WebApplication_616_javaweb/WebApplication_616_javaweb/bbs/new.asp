<%save = true
function zweistellig(wert)
if len(wert)<2 then wert = "0" & wert
zweistellig = wert
end function
function check1(wert)
	if len(wert)<1 then save = false
	wert = replace(wert,"'", "`")
	wert = replace(wert,chr(34), "``")
	check1 = wert
end function
function check2(wert)
	if len(wert)<1 then wert = " "
	wert = replace(wert,"'", "`")
	wert = replace(wert,chr(34), "``")
	check2 = wert
end function
autor = check1(request.form("autor"))
title = check1(request.form("title"))
message = check1(request.form("message"))
mail = check2(request.form("mail"))
if save = true then
	updated = zweistellig(right(year(date),2)) & zweistellig(month(date)) & zweistellig(day(date)) & zweistellig(hour(time)) & zweistellig(minute(time)) & zweistellig(second(time))
	set db = Server.CreateObject("ADODB.Connection")
	connect="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("data.asp")
	db.Open connect
	sql = "insert into postings (title, autor, mail, message, views, replies, datum, datum_updated, updated, connected)"
	sql = sql & " values('"&title&"', '"&autor&"', '"&mail&"','"&message&"', 0, 0, '"& date & " " & time &"', '"& date & " " & time &"', '"&updated&"', 0)" 
	db.Execute(sql)
	db.close
	set db = nothing
	response.redirect("index.asp")
else%>
<script language="vbscript">
		msgbox "请检查填写项目是否完整！",vbInformation,"出错提示！"
		window.location="vbscript:history.back"
</script>
<%end if%>
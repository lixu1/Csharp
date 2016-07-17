<%
'文件属性：例如上传文件为c:\myfile\doc.txt
'FileName    文件名       字符串    "doc.txt"
'FileSize    文件大小     数值       1210
'FileType    文件类型     字符串    "text/plain"
'FileExt     文件扩展名   字符串    "txt"
'FilePath    文件原路径   字符串    "c:\myfile"

Const MaxFileSize = 300000 '上传文件大小限制
Const SaveUpFilesPath = "uploadfile" '存放上传文件的目录
Const UpFileType = "gif|jpg|bmp|png" '允许的上传文件类型
Dim oUpFileStream

Class upload_file
    Dim Form, File, Version

    Private Sub Class_Initialize
        '定义变量
        Dim RequestBinDate, sStart, bCrLf, sInfo, iInfoStart, iInfoEnd, tStream, iStart, oFileInfo
        Dim iFileSize, sFilePath, sFileType, sFormvalue, sFileName
        Dim iFindStart, iFindEnd
        Dim iFormStart, iFormEnd, sFormName
        '代码开始
        Version = "无组件上传类 Version 0.96"
        Set Form = Server.CreateObject("Scripting.Dictionary")
        Set File = Server.CreateObject("Scripting.Dictionary")
        If Request.TotalBytes < 1 Then Exit Sub
        Set tStream = Server.CreateObject("adodb.stream")
        Set oUpFileStream = Server.CreateObject("adodb.stream")
        oUpFileStream.Type = 1
        oUpFileStream.Mode = 3
        oUpFileStream.Open
        oUpFileStream.Write Request.BinaryRead(Request.TotalBytes)
        oUpFileStream.Position = 0
        RequestBinDate = oUpFileStream.Read
        iFormEnd = oUpFileStream.Size
        bCrLf = chrB(13) & chrB(10)
        '取得每个项目之间的分隔符
        sStart = MidB(RequestBinDate, 1, InStrB(1, RequestBinDate, bCrLf) -1)
        iStart = LenB (sStart)
        iFormStart = iStart + 2
        '分解项目
        Do
            iInfoEnd = InStrB(iFormStart, RequestBinDate, bCrLf & bCrLf) + 3
            tStream.Type = 1
            tStream.Mode = 3
            tStream.Open
            oUpFileStream.Position = iFormStart
            oUpFileStream.CopyTo tStream, iInfoEnd - iFormStart
            tStream.Position = 0
            tStream.Type = 2
            tStream.Charset = "gb2312"
            sInfo = tStream.ReadText
            '取得表单项目名称
            iFormStart = InStrB(iInfoEnd, RequestBinDate, sStart) -1
            iFindStart = InStr(22, sInfo, "name=""", 1) + 6
            iFindEnd = InStr(iFindStart, sInfo, """", 1)
            sFormName = Mid (sinfo, iFindStart, iFindEnd - iFindStart)
            '如果是文件
            If InStr (45, sInfo, "filename=""", 1) > 0 Then
                Set oFileInfo = New FileInfo
                '取得文件属性
                iFindStart = InStr(iFindEnd, sInfo, "filename=""", 1) + 10
                iFindEnd = InStr(iFindStart, sInfo, """", 1)
                sFileName = Mid (sinfo, iFindStart, iFindEnd - iFindStart)
                oFileInfo.FileName = GetFileName(sFileName)
                oFileInfo.FilePath = GetFilePath(sFileName)
                oFileInfo.FileExt = GetFileExt(sFileName)
                iFindStart = InStr(iFindEnd, sInfo, "Content-Type: ", 1) + 14
                iFindEnd = InStr(iFindStart, sInfo, vbCr)
                oFileInfo.FileType = Mid (sinfo, iFindStart, iFindEnd - iFindStart)
                oFileInfo.FileStart = iInfoEnd
                oFileInfo.FileSize = iFormStart - iInfoEnd -2
                oFileInfo.FormName = sFormName
                File.Add sFormName, oFileInfo
            Else
                '如果是表单项目
                tStream.Close
                tStream.Type = 1
                tStream.Mode = 3
                tStream.Open
                oUpFileStream.Position = iInfoEnd
                oUpFileStream.CopyTo tStream, iFormStart - iInfoEnd -2
                tStream.Position = 0
                tStream.Type = 2
                tStream.Charset = "gb2312"
                sFormvalue = tStream.ReadText
                Form.Add sFormName, sFormvalue
            End If
            tStream.Close
            iFormStart = iFormStart + iStart + 2
            '如果到文件尾了就退出
        Loop Until (iFormStart + 2) = iFormEnd
        RequestBinDate = ""
        Set tStream = Nothing
    End Sub

    Private Sub Class_Terminate
        '清除变量及对像
        If Not Request.TotalBytes<1 Then
            oUpFileStream.Close
            Set oUpFileStream = Nothing
        End If
        Form.RemoveAll
        File.RemoveAll
        Set Form = Nothing
        Set File = Nothing
    End Sub

    '取得文件路径

    Private Function GetFilePath(FullPath)
        If FullPath <> "" Then
            GetFilePath = Left(FullPath, InStrRev(FullPath, "\"))
        Else
            GetFilePath = ""
        End If
    End Function

    '取得文件名

    Private Function GetFileName(FullPath)
        If FullPath <> "" Then
            GetFileName = Mid(FullPath, InStrRev(FullPath, "\") + 1)
        Else
            GetFileName = ""
        End If
    End Function

    '取得扩展名

    Private Function GetFileExt(FullPath)
        If FullPath <> "" Then
            GetFileExt = Mid(FullPath, InStrRev(FullPath, ".") + 1)
        Else
            GetFileExt = ""
        End If
    End Function

End Class

'文件属性类

Class FileInfo
    Dim FormName, FileName, FilePath, FileSize, FileType, FileStart, FileExt

    Private Sub Class_Initialize
        FileName = ""
        FilePath = ""
        FileSize = 0
        FileStart = 0
        FormName = ""
        FileType = ""
        FileExt = ""
    End Sub

    '保存文件方法

    Public Function SaveToFile(FullPath)
        Dim oFileStream, ErrorChar, i
        SaveToFile = 1
        If Trim(fullpath) = "" Or Right(fullpath, 1) = "/" Then Exit Function
        Set oFileStream = CreateObject("Adodb.Stream")
        oFileStream.Type = 1
        oFileStream.Mode = 3
        oFileStream.Open
        oUpFileStream.position = FileStart
        oUpFileStream.copyto oFileStream, FileSize
        oFileStream.SaveToFile FullPath, 2
        oFileStream.Close
        Set oFileStream = Nothing
        SaveToFile = 0
    End Function

End Class

Const upload_type = 0 '上传方法：0=无惧无组件上传类，1=FSO上传 2=lyfupload，3=aspupload，4=chinaaspupload

Dim upload, File, formName, SavePath, filename, fileExt
Dim upNum
Dim EnableUpload
Dim Forumupload
Dim ranNum
Dim uploadfiletype
Dim msg, founderr
msg = ""
founderr = False
EnableUpload = False
SavePath = SaveUpFilesPath '存放上传文件的目录
If Right(SavePath, 1)<>"/" Then SavePath = SavePath&"/" '在目录后加(/)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<%
If EnableUploadFile = "No" Then
    response.Write "系统未开放文件上传功能"
Else
    If 1=0 Then
        response.Write("请登录后再使用本功能！")
    Else
        Select Case upload_type
            Case 0
                Call upload_0() '使用化境无组件上传类
            Case Else
                'response.write "本系统未开放插件功能"
                'response.end
        End Select
    End If
End If
%>
</body>
</html>
<%
Sub upload_0() '使用化境无组件上传类
    Set upload = New upload_file '建立上传对象
    For Each formName in upload.File '列出所有上传了的文件
        Set File = upload.File(formName) '生成一个文件对象
        If File.filesize<100 Then
            msg = "请先选择你要上传的文件！"
            founderr = True
        End If
        If File.filesize>(MaxFileSize * 1024) Then
            msg = "文件大小超过了限制，最大只能上传" & CStr(MaxFileSize) & "K的文件！"
            founderr = True
        End If

        fileExt = LCase(File.FileExt)
        Forumupload = Split(UpFileType, "|")
        For i = 0 To UBound(Forumupload)
            If fileEXT = Trim(Forumupload(i)) Then
                EnableUpload = True
                Exit For
            End If
        Next
        If fileEXT = "asp" Or fileEXT = "asa" Or fileEXT = "aspx" Then
            EnableUpload = False
        End If
        If EnableUpload = False Then
            msg = "这种文件类型不允许上传！\n\n只允许上传这几种文件类型：" & UpFileType
            founderr = True
        End If

        strJS = "<SCRIPT language=javascript>" & vbCrLf
        If founderr<>True Then
            Randomize
            ranNum = Int(900 * Rnd) + 100
            filename = SavePath&Year(Now)&Month(Now)&Day(Now)&Hour(Now)&Minute(Now)&Second(Now)&ranNum&"."&fileExt
            File.SaveToFile Server.mappath(FileName) '保存文件
            msg = "上传图片成功！"
%>
<script language='javascript'>
parent.document.form1.imgpath.value="<%=filename%>"
alert('<%=msg%>');
//alert('上传文件成功');
 history.go(-1);
</script>
<%
End If
Set File = Nothing
Next
Set upload = Nothing
End Sub
%>

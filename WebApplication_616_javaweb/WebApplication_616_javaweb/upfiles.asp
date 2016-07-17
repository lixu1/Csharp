<%
'�ļ����ԣ������ϴ��ļ�Ϊc:\myfile\doc.txt
'FileName    �ļ���       �ַ���    "doc.txt"
'FileSize    �ļ���С     ��ֵ       1210
'FileType    �ļ�����     �ַ���    "text/plain"
'FileExt     �ļ���չ��   �ַ���    "txt"
'FilePath    �ļ�ԭ·��   �ַ���    "c:\myfile"

Const MaxFileSize = 300000 '�ϴ��ļ���С����
Const SaveUpFilesPath = "uploadfile" '����ϴ��ļ���Ŀ¼
Const UpFileType = "gif|jpg|bmp|png" '������ϴ��ļ�����
Dim oUpFileStream

Class upload_file
    Dim Form, File, Version

    Private Sub Class_Initialize
        '�������
        Dim RequestBinDate, sStart, bCrLf, sInfo, iInfoStart, iInfoEnd, tStream, iStart, oFileInfo
        Dim iFileSize, sFilePath, sFileType, sFormvalue, sFileName
        Dim iFindStart, iFindEnd
        Dim iFormStart, iFormEnd, sFormName
        '���뿪ʼ
        Version = "������ϴ��� Version 0.96"
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
        'ȡ��ÿ����Ŀ֮��ķָ���
        sStart = MidB(RequestBinDate, 1, InStrB(1, RequestBinDate, bCrLf) -1)
        iStart = LenB (sStart)
        iFormStart = iStart + 2
        '�ֽ���Ŀ
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
            'ȡ�ñ���Ŀ����
            iFormStart = InStrB(iInfoEnd, RequestBinDate, sStart) -1
            iFindStart = InStr(22, sInfo, "name=""", 1) + 6
            iFindEnd = InStr(iFindStart, sInfo, """", 1)
            sFormName = Mid (sinfo, iFindStart, iFindEnd - iFindStart)
            '������ļ�
            If InStr (45, sInfo, "filename=""", 1) > 0 Then
                Set oFileInfo = New FileInfo
                'ȡ���ļ�����
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
                '����Ǳ���Ŀ
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
            '������ļ�β�˾��˳�
        Loop Until (iFormStart + 2) = iFormEnd
        RequestBinDate = ""
        Set tStream = Nothing
    End Sub

    Private Sub Class_Terminate
        '�������������
        If Not Request.TotalBytes<1 Then
            oUpFileStream.Close
            Set oUpFileStream = Nothing
        End If
        Form.RemoveAll
        File.RemoveAll
        Set Form = Nothing
        Set File = Nothing
    End Sub

    'ȡ���ļ�·��

    Private Function GetFilePath(FullPath)
        If FullPath <> "" Then
            GetFilePath = Left(FullPath, InStrRev(FullPath, "\"))
        Else
            GetFilePath = ""
        End If
    End Function

    'ȡ���ļ���

    Private Function GetFileName(FullPath)
        If FullPath <> "" Then
            GetFileName = Mid(FullPath, InStrRev(FullPath, "\") + 1)
        Else
            GetFileName = ""
        End If
    End Function

    'ȡ����չ��

    Private Function GetFileExt(FullPath)
        If FullPath <> "" Then
            GetFileExt = Mid(FullPath, InStrRev(FullPath, ".") + 1)
        Else
            GetFileExt = ""
        End If
    End Function

End Class

'�ļ�������

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

    '�����ļ�����

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

Const upload_type = 0 '�ϴ�������0=�޾�������ϴ��࣬1=FSO�ϴ� 2=lyfupload��3=aspupload��4=chinaaspupload

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
SavePath = SaveUpFilesPath '����ϴ��ļ���Ŀ¼
If Right(SavePath, 1)<>"/" Then SavePath = SavePath&"/" '��Ŀ¼���(/)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<%
If EnableUploadFile = "No" Then
    response.Write "ϵͳδ�����ļ��ϴ�����"
Else
    If 1=0 Then
        response.Write("���¼����ʹ�ñ����ܣ�")
    Else
        Select Case upload_type
            Case 0
                Call upload_0() 'ʹ�û���������ϴ���
            Case Else
                'response.write "��ϵͳδ���Ų������"
                'response.end
        End Select
    End If
End If
%>
</body>
</html>
<%
Sub upload_0() 'ʹ�û���������ϴ���
    Set upload = New upload_file '�����ϴ�����
    For Each formName in upload.File '�г������ϴ��˵��ļ�
        Set File = upload.File(formName) '����һ���ļ�����
        If File.filesize<100 Then
            msg = "����ѡ����Ҫ�ϴ����ļ���"
            founderr = True
        End If
        If File.filesize>(MaxFileSize * 1024) Then
            msg = "�ļ���С���������ƣ����ֻ���ϴ�" & CStr(MaxFileSize) & "K���ļ���"
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
            msg = "�����ļ����Ͳ������ϴ���\n\nֻ�����ϴ��⼸���ļ����ͣ�" & UpFileType
            founderr = True
        End If

        strJS = "<SCRIPT language=javascript>" & vbCrLf
        If founderr<>True Then
            Randomize
            ranNum = Int(900 * Rnd) + 100
            filename = SavePath&Year(Now)&Month(Now)&Day(Now)&Hour(Now)&Minute(Now)&Second(Now)&ranNum&"."&fileExt
            File.SaveToFile Server.mappath(FileName) '�����ļ�
            msg = "�ϴ�ͼƬ�ɹ���"
%>
<script language='javascript'>
parent.document.form1.imgpath.value="<%=filename%>"
alert('<%=msg%>');
//alert('�ϴ��ļ��ɹ�');
 history.go(-1);
</script>
<%
End If
Set File = Nothing
Next
Set upload = Nothing
End Sub
%>

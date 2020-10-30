<%
'@title: Class_Fso
'@author: ekede.com
'@date: 2017-02-13
'@description: FileSystemObject,Stream文件系统操作类

Class Class_Fso
    '
    Private objFSO

    Private Sub Class_Initialize()
        Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    End Sub

    Private Sub Class_Terminate()
        Set objFSO = Nothing
    End Sub

    '@GetMapPath(ByRef path): 获取物理路径

    Public Function GetMapPath(ByRef path)
        If StrCheck(path) Then
            GetMapPath = -1
        Else
            GetMapPath = server.mappath(path)
        End If
    End Function

    '@GetRealPath(ByRef path): 相对路径判断文件是否存在,并返回物理路径

    Public Function GetRealPath(ByRef path)
        Dim msg, real
        real = GetMapPath(path)
        If real = -1 Then
            GetRealPath = -1
        Else
            If objFSO.FileExists(real) Then
                GetRealPath = real
            Else
                GetRealPath = -1
            End If
        End If
    End Function
	
	'判断是否包含路径非法字符
	Private Function StrCheck(ByRef str)
		Dim i, arrays
		StrCheck = False
		If IsNull(str) Or Trim(str) = Empty Then Exit Function
		'
	    arrays = Split(":,*,?,"",<,>,|",",")
		For i = 0 To UBound(arrays)
			If InStr(str, arrays(i)) > 0 Then
				StrCheck = True
				Exit Function
			End If
		Next
	End Function

    '=======文件操作========
	
    '@ReportFileStatus(ByRef fileName): 文件是否存在？

    Public Function ReportFileStatus(ByRef fileName)
        Dim msg
        msg = -1
        If (objFSO.FileExists(fileName)) Then
            msg = 1
        Else
            msg = -1
        End If
        ReportFileStatus = msg
    End Function
	
    '@GetFileObject(ByRef fileName): 文件转换为对象

    Public Function GetFileObject(ByRef fileName)
        Set GetFileObject = objFSO.GetFile(fileName)
    End Function

    '@DeleteAFile(ByRef fileSpec): 文件删除

    Public Function DeleteAFile(ByRef fileSpec)
        If ReportFileStatus(FileSpec) = 1 Then
            objFSO.DeleteFile(fileSpec)
            DeleteAFile = 1
        Else
            DeleteAFile = -1
        End If
    End Function

    '@CopyAFile(ByRef sourceFile,ByRef destinationFile): 文件复制

    Public Function CopyAFile(ByRef sourceFile,ByRef destinationFile)
        Dim MyFile
        If ReportFileStatus(sourceFile) = 1 Then
            Set MyFile = objFSO.GetFile(sourceFile)
            MyFile.Copy (destinationFile)
            CopyAFile = 1
        Else
            CopyAFile = -1
        End If
    End Function

    '@MoveAFile(ByRef sourceFile,ByRef destinationFileOrPath): 文件移动

    Public Function MoveAFile(ByRef sourceFile,ByRef destinationFileOrPath)
        If ReportFileStatus(sourceFile) = 1 And ReportFileStatus(destinationFileOrPath) = -1 Then
            objFSO.MoveFile sourceFile, destinationFileOrPath
            MoveAFile = 1
        Else
            MoveAFile = -1
        End If
    End Function
	
    '@GetFileSize(ByRef fileName): 取文件大小

    Public Function GetFileSize(ByRef fileName)
        Dim f
        If ReportFileStatus(fileName) = 1 Then
            Set f = objFSO.GetFile(fileName)
            GetFileSize = f.Size
        Else
            GetFileSize = -1
        End If
    End Function

    '@ShowDatecreated(ByRef fileSpec): 文件创建日期

    Public Function ShowDatecreated(ByRef fileSpec)
        Dim f
        If ReportFileStatus(fileSpec) = 1 Then
            Set f = objFSO.GetFile(fileSpec)
            ShowDatecreated = f.DateCreated
        Else
            ShowDatecreated = -1
        End If
    End Function

    '@GetAttributes(ByRef fileSpec): 文件属性

    Public Function GetAttributes(ByRef fileName)
        Dim f
        Dim strFileAttributes
        If ReportFileStatus(fileName) = 1 Then
            Set f = objFSO.GetFile(fileName)
            Select Case f.Attributes
                Case 0 strFileAttributes = "普通文件。没有设置任何属性。 "
                Case 1 strFileAttributes = "只读文件。可读写。 "
                Case 2 strFileAttributes = "隐藏文件。可读写。 "
                Case 4 strFileAttributes = "系统文件。可读写。 "
                Case 16 strFileAttributes = "文件夹或目录。只读。 "
                Case 32 strFileAttributes = "上次备份后已更改的文件。可读写。 "
                Case 1024 strFileAttributes = "链接或快捷方式。只读。 "
                Case 2048 strFileAttributes = " 压缩文件。只读。"
            End Select
            GetAttributes = strFileAttributes
        Else
            GetAttributes = -1
        End If
    End Function

    '@ShowFileAccessInfo(ByRef fileName,ByRef infoType): 显示文件创建时信息

    Public Function ShowFileAccessInfo(ByRef fileName,ByRef infoType)
        '// 1 -----创建时间
        '// 2 -----上次访问时间
        '// 3 -----上次修改时间
        '// 4 -----文件路径
        '// 5 -----文件名称
        '// 6 -----文件类型
        '// 7 -----文件大小
        '// 8 -----父目录
        '// 9 -----根目录
        '// 10 -----文件属性
        Dim f, s
        If ReportFileStatus(fileName) = 1 Then
            Set f = objFSO.GetFile(fileName)
            Select Case infoType
                Case 1 s = f.DateCreated
                Case 2 s = f.DateLastAccessed
                Case 3 s = f.DateLastModified
                Case 4 s = f.Path
                Case 5 s = f.Name
                Case 6 s = f.Type
                Case 7 s = f.Size
                Case 8 s = f.ParentFolder
                Case 9 s = f.RootFolder
                Case 10 s = f.Attributes
            End Select
            ShowFileAccessInfo = s
        Else
            ShowFileAccessInfo = -1
        End If
    End Function
	
    '=======文本文件操作========

    '@CreateTxtFile(ByRef fileName,ByRef textStr): 文本文件创建

    Public Function CreateTxtFile(ByRef fileName,ByRef textStr)
        Dim f
        If ReportFileStatus(fileName) = 1 Then
            CreateTxtFile = -1
        Else
            Set f = objFSO.CreateTextFile(fileName, true, false)
            If textStr<> "" Then f.WriteLine textStr
            CreateTxtFile = 1
        End If
    End Function

    '@WriteTxtFile(ByRef fileName,ByRef textStr,ByRef writeORAppendType): 写文本文件

    Public Function WriteTxtFile(ByRef fileName,ByRef textStr,ByRef writeORAppendType)
        Const ForReading = 1, ForWriting = 2 , ForAppending = 8
        Dim f, m
        Select Case writeORAppendType
		Case 1: '文件进行写操作
				Set f = objFSO.OpenTextFile(fileName, ForWriting, True)
				f.Write textStr
				f.Close
				If ReportFileStatus(FileName) = 1 Then
					WriteTxtFile = 1
				Else
					WriteTxtFile = -1
				End If
		Case 2: '文件末尾进行写操作
				If ReportFileStatus(fileName) = 1 Then
					Set f = objFSO.OpenTextFile(fileName, ForAppending)
					f.Write textStr
					f.Close
					WriteTxtFile = 1
				Else
					WriteTxtFile = -1
				End If
		Case 3: '文件末尾进行写操作 换行 不存在创建文件
				If ReportFileStatus(fileName) = 1 Then
					Set f = objFSO.OpenTextFile(fileName, ForAppending)
					f.WriteLine textStr
					f.Close
					WriteTxtFile = 1
				Else
					WriteTxtFile = CreateTxtFile(fileName, textStr)
				End If
        End Select
    End Function

    '@ReadTxtFile(ByRef fileName): 读文本文件

    Public Function ReadTxtFile(ByRef fileName)
        Const ForReading = 1, ForWriting = 2
        Dim f, m
        If ReportFileStatus(fileName) = 1 Then
            Set f = objFSO.OpenTextFile(fileName, ForReading)
            m = f.ReadAll 'ReadLine
            ReadTxtFile = m
            f.Close
        Else
            ReadTxtFile = -1
        End If
    End Function

    '=======目录操作========
	
    '@ReportFolderStatus(ByRef folder): 判断目录是否存在

    Public Function ReportFolderStatus(ByRef folder)
        Dim msg
        msg = -1
        If (objFSO.FolderExists(folder)) Then
            msg = 1
        Else
            msg = -1
        End If
        ReportFolderStatus = msg
    End Function
	
    '@GetFolderObject(ByRef folder): 目录转换为对象

    Public Function GetFolderObject(ByRef folder)
        Set GetFolderObject = objFSO.GetFolder(folder)
    End Function

    '@GetFolderSize(ByRef folderName): 取目录大小

    Public Function GetFolderSize(ByRef folderName)
        Dim f
        If ReportFolderStatus(folderName) = 1 Then
            Set f = objFSO.GetFolder(folderName)
            GetFolderSize = f.Size
        Else
            GetFolderSize = -1
        End If
    End Function

    '@CreateAFolder(ByRef folderSpec): 创建的文件夹

    Public Function CreateAFolder(ByRef folderSpec)
        On Error Resume Next
        Dim f
        If ReportFolderStatus(folderSpec) = 1 Then
            CreateAFolder = -1
        Else
            Set f = objFSO.CreateFolder(folderSpec)
            CreateAFolder = 1
        End If
    End Function

    '@CreateFolders(ByRef folderSpec): 创建多级文件夹

    Public Function CreateFolders(ByRef folderSpec)
        Dim f
        If ReportFolderStatus(folderSpec) = 1 Then
            CreateFolders = -1
        Else
            Dim astrPath, ulngPath, strTmpPath
            astrPath = Split(folderSpec, "\")
            ulngPath = UBound(astrPath)
            strTmpPath = ""
            For i = 0 To ulngPath
                strTmpPath = strTmpPath & astrPath(i) & "\"
                CreateAFolder(strTmpPath)
            Next
            CreateFolders = 1
        End If
    End Function

    '@DeleteAFolder(ByRef folderSpec): 目录删除

    Public Function DeleteAFolder(ByRef folderSpec)
        If ReportFolderStatus(folderSpec) = 1 Then
            objFSO.DeleteFolder (folderSpec)
            DeleteAFolder = 1
        Else
            DeleteAFolder = -1
        End If
    End Function

    '@ShowFolderList(ByRef folderSpec): 目录列表

    Public Function ShowFolderList(ByRef folderSpec)
        Dim f, f1, fc, s, i
        If ReportFolderStatus(folderSpec) = 1 Then
            Set f = objFSO.GetFolder(folderSpec)
            Set fc = f.SubFolders
			i=0
            For Each f1 in fc
			    If i = 0 Then
	               s = s & f1.Name
				Else
	               s = s & "|" & f1.Name
				End If
				i = i + 1
            Next
            ShowFolderList = s
        Else
            ShowFolderList = -1
        End If
    End Function
	
    '@ShowFileList(ByRef folderSpec): 显示文件列表

    Public Function ShowFileList(ByRef folderSpec)
        Dim f, f1, fc, s
        If ReportFolderStatus(folderSpec) = 1 Then
            Set f = objFSO.GetFolder(folderSpec)
            Set fc = f.Files
			i=0
            For Each f1 in fc
			    If i = 0 Then
	               s = s & f1.Name
				Else
	               s = s & "|" & f1.Name
				End If
				i = i + 1
            Next
            ShowFileList = s
        Else
            ShowFileList = -1
        End If
    End Function

    '@CopyAFolder(ByRef sourceFolder,ByRef destinationFolder): 目录复制

    Public Function CopyAFolder(ByRef sourceFolder,ByRef destinationFolder)
        objFSO.CopyFolder sourceFolder, destinationFolder
        CopyAFolder = 1
        CopyAFolder = -1
    End Function

    '@MoveAFolder(ByRef sourcePath,ByRef destinationPath): 目录进行移动

    Public Function MoveAFolder(ByRef sourcePath,ByRef destinationPath)
        If ReportFolderStatus(sourcePath) = 1 And ReportFolderStatus(destinationPath) = 0 Then
            objFSO.MoveFolder sourcePath, destinationPath
            MoveAFolder = 1
        Else
            MoveAFolder = -1
        End If
    End Function

    '@ShowFolderAccessInfo(ByRef folderName,ByRef infoType): 目录时间,名称,大小,类型,父目录,根目录

    Public Function ShowFolderAccessInfo(ByRef folderName,ByRef infoType)
        '//功能：显示目录创建时信息
        '//形参：目录名,信息类别
        '// 1 -----创建时间
        '// 2 -----上次访问时间
        '// 3 -----上次修改时间
        '// 4 -----目录路径
        '// 5 -----目录名称
        '// 6 -----目录类型
        '// 7 -----目录大小
        '// 8 -----父目录
        '// 9 -----根目录
        Dim f, s
        If ReportFolderStatus(folderName) = 1 Then
            Set f = objFSO.GetFolder(folderName)
            Select Case infoType
                Case 1 s = f.DateCreated
                Case 2 s = f.DateLastAccessed
                Case 3 s = f.DateLastModified
                Case 4 s = f.Path
                Case 5 s = f.Name
                Case 6 s = f.Type
                Case 7 s = f.Size
                Case 8 s = f.ParentFolder
                Case 9 s = f.IsRootFolder
            End Select
            ShowFolderAccessInfo = s
        Else
            ShowFolderAccessInfo = -1
        End If
    End Function

    '@DisplayLevelDepth(ByRef pathSpec): 遍历目录

    Public Function DisplayLevelDepth(ByRef folderSpec)
        Dim f, n , path
        If  ReportFolderStatus(folderSpec) = 1 Then
            Set f = objFSO.GetFolder(folderSpec)
			If f.IsRootFolder Then
				'DisplayLevelDepth = "指定的文件夹是根文件夹。"&RootFolder
				DisplayLevelDepth = 1
			Else
				Do Until f.IsRootFolder
					path = path & f.Name &"<br>"
					Set f = f.ParentFolder
					n = n + 1
				Loop
				'DisplayLevelDepth = "指定的文件夹是嵌套级为 " & n & " 的文件夹。<br>" & path
				DisplayLevelDepth = n
			End If
		Else
			DisplayLevelDepth = - 1
		End If
    End Function

    '========磁盘操作========

    '@ReportDriveStatus(ByRef drv): 驱动器是否存在？

    Public Function ReportDriveStatus(ByRef drv)
        Dim msg
        msg = -1
        If objFSO.DriveExists(drv) Then
            msg = 1
        Else
            msg = -1
        End If
        ReportDriveStatus = msg
    End Function

    '@ShowFileSystemType(ByRef drvspec): 可用的返回类型包括 FAT、NTFS 和 CDFS。

    Public Function ShowFileSystemType(ByRef drvspec)
        Dim d
        If ReportDriveStatus(drvspec) = 1 Then
            Set d = objFSO.GetDrive(drvspec)
            ShowFileSystemType = d.FileSystem
        Else
            ShowFileSystemType = -1
        End If
    End Function
	
    '=======Stream操作========

    '@Reads(ByRef fileName,ByRef cset): Stream 读文件 文本 cset空为二进制

    Public Function Reads(ByRef fileName,ByRef cset)
        If ReportFileStatus(fileName) = 1 Then
            Set objStream = Server.CreateObject("ADODB.Stream")
            ObjStream.Type = 1
            ObjStream.Mode = 3
            ObjStream.Open
            ObjStream.LoadFromFile fileName
            ObjStream.Position = 0
			If cset = "" Then
               Reads = ObjStream.Read()
			Else
			   objStream.Type = 2
			   objStream.Charset = cset
			   Reads = objStream.ReadText()
			End If
			objStream.Close
		    set objStream = nothing
        Else
            Reads = -1
        End If
    End Function

    '@Writes(ByRef fileName,ByRef content,ByRef cset): Stream 写文件 cset空为二进制

    Public Function Writes(fileName, content, cset)
        On Error Resume Next
		dim objStream
        Set objStream = Server.CreateObject("ADODB.Stream")
		If  cset = "" Then
			ObjStream.Type = 1 '二进制
			ObjStream.Mode = 3 '1读, 2写, 3读写
			ObjStream.Open
			ObjStream.Position = 0
			ObjStream.Write Content
        Else
			ObjStream.Type = 2 '文本
			ObjStream.Mode = 3
			ObjStream.Open
			objStream.Position = 0
			ObjStream.Charset = cset
			ObjStream.WriteText = Content
		End If
        ObjStream.SaveToFile fileName, 2 '1创建，2覆盖
        ObjStream.Close
        Set objStream = Nothing
        If Err Then
            Err.Clear
            Writes = -1
        Else
            Writes = 1
        End If
    End Function

	'@Iconv(ByRef inCset,ByRef OutCset,ByRef content): Stream 编码转换UTF-8,GB2312

	Public Function Iconv(ByRef inCset,ByRef OutCset,ByRef content)
		Dim objStream
		Set objStream = server.CreateObject("Adodb.Stream")
			objStream.Charset = inCset
			objStream.Type = 2
			objStream.Mode = 3
			objStream.Open
			objStream.Position = 0
			objStream.WriteText content
			objStream.Position = 0
			objStream.Charset = OutCset
			Iconv = objStream.ReadText
			objStream.Close
		set objStream = nothing
	End Function
	
	'@Bytes2Str(ByRef body,ByRef cset): Stream 字节流转字符串 body字节数组,cset编码格式
	
	Function Bytes2Str(ByRef body,ByRef cset)
		dim objStream
		set objStream = Server.CreateObject("adodb.stream")
			objStream.Type = 1 '以二进制模式打开
			objStream.Mode =3
			objStream.Open
			objStream.Position = 0
			objStream.Write body
			objStream.Position = 0
			objStream.Type = 2
			objStream.Charset = cset
			Bytes2Str = objStream.ReadText
		    objStream.Close
		set objStream = nothing
	End Function
	
	'@Str2Bytes(ByRef text,ByRef cset): Stream 字符串转字节流 Text字符窜,cset编码格式
	
	Function Str2Bytes(ByRef text,ByRef cset)
		dim objStream
		set objStream = Server.CreateObject("adodb.stream")
			objStream.Charset = cset
			objStream.Type = 2 '以文本模式打开
			objStream.Mode =3
			objStream.Open
			objStream.Position = 0
			objStream.WriteText text
			objStream.Position = 0
			objStream.Type = 1
			Str2Bytes = objStream.Read
		    objStream.Close
		set objStream = nothing
	End Function

End Class
%>
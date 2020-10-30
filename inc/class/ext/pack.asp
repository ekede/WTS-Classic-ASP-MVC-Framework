<%
'@title: Class_Ext_Pack
'@author: ekede.com
'@date: 2017-02-13
'@description: 打包解包文件夹

Class Class_Ext_Pack

    Private pathDir_

    '@Pack(ByRef pathDir,ByRef pathFile): 将目录pathDir打包成pathFile

    Public Sub Pack(ByRef pathDir,ByRef pathFile)
	   '创建一个空的XML文件，为写入文件作准备
        Dim XmlDoc, Root
        Set XmlDoc = Server.CreateObject("Microsoft.XMLDOM")
			XmlDoc.async = False
			Set Root = XmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
				XmlDoc.appendChild(Root)
				XmlDoc.appendChild(XmlDoc.CreateElement("root"))
				XmlDoc.Save(pathFile)
			Set Root = Nothing
        Set XmlDoc = Nothing
        '格式化路径
		If Right(pathDir,1)<>"\" Then pathDir = pathDir&"\"
		pathDir_ = pathDir
		'递归加载文件到xml
		LoadData pathDir, pathFile
    End Sub

    '遍历目录内的所有文件以及文件夹

    Private Sub LoadData(ByRef pathDir,ByRef pathFile)
        Dim XmlDoc
        Dim fso 'fso对象
        Dim objFolder '文件夹对象
        Dim objSubFolders '子文件夹集合
        Dim objSubFolder '子文件夹对象
        Dim objFiles '文件集合
        Dim objFile '文件对象
        Dim objStream
        Dim pathname, TextStream, pp, Xfolder, Xfpath, Xfile, Xpath, Xstream
        Dim PathNameStr
        '
        Set fso = server.CreateObject("scripting.filesystemobject")
        Set objFolder = fso.GetFolder(pathDir)'创建文件夹对象
        '
        Set XmlDoc = Server.CreateObject("Microsoft.XMLDOM")
        XmlDoc.load pathFile
        XmlDoc.async = False
        '写入每个文件夹路径
        Set Xfolder = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("folder"))
        Set Xfpath = Xfolder.AppendChild(XmlDoc.CreateElement("path"))
        Xfpath.text = Replace(pathDir, pathDir_, "")
        Set objFiles = objFolder.Files
        For Each objFile in objFiles
            If LCase(pathDir & objFile.Name)<> LCase(Request.ServerVariables("PATH_TRANSLATED"))Then
                PathNameStr = pathDir & objFile.Name
                '写入文件的路径及文件内容
                Set Xfile = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("file"))
                Set Xpath = Xfile.AppendChild(XmlDoc.CreateElement("path"))
                Xpath.text = Replace(PathNameStr, pathDir_, "")
                '读文件流
                Set objStream = Server.CreateObject("ADODB.Stream")
                objStream.Type = 1
                objStream.Open()
                objStream.LoadFromFile(PathNameStr)
                objStream.position = 0
                '流转base64
                Set Xstream = Xfile.AppendChild(XmlDoc.CreateElement("stream"))
                Xstream.SetAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes"
                Xstream.dataType = "bin.base64"
                Xstream.nodeTypedValue = objStream.Read()
                Set Xstream = Nothing
                Set objStream = Nothing
                Set Xpath = Nothing
                Set Xfile = Nothing
            End If
        Next
        XmlDoc.Save(pathFile)
        Set Xfpath = Nothing
        Set Xfolder = Nothing
        Set XmlDoc = Nothing

        '创建的子文件夹对象	 调用递归遍历子文件夹
        Set objSubFolders = objFolder.SubFolders
        For Each objSubFolder in objSubFolders
            pathName = pathDir & objSubFolder.Name &"\"
            Call LoadData(pathName, pathFile)
        Next
        Set objSubFolders = Nothing
        '
        Set objFolder = Nothing
        Set fso = Nothing
    End Sub

    '@UnPack(ByRef pathFile,ByRef pathDir): 将pathFile解包到pathDir

    Public Sub UnPack(ByRef pathFile,ByRef pathDir)
        On Error Resume Next
        Dim objXmlFile
        Dim objNodeList
        Dim objFSO
        Dim objStream
        Dim i, j
		If Right(pathDir,1)<>"\" Then pathDir = pathDir&"\"
        '
        Set objXmlFile = Server.CreateObject("Microsoft.XMLDOM")
        objXmlFile.load(pathFile)

        If objXmlFile.readyState = 4 Then
            If objXmlFile.parseError.errorCode = 0Then
                '输出目录
                Set objNodeList = objXmlFile.documentElement.selectNodes("//folder/path")
                Set objFSO = CreateObject("Scripting.FileSystemObject")
                j = objNodeList.Length -1
                For i = 0 To j
                    If objFSO.FolderExists(pathDir & objNodeList(i).text) = False Then
                        objFSO.CreateFolder(pathDir & objNodeList(i).text)
                    End If
                Next
                Set objFSO = Nothing
                Set objNodeList = Nothing
                '输出文件
                Set objNodeList = objXmlFile.documentElement.selectNodes("//file/path")
                j = objNodeList.Length -1
                For i = 0 To j
                    Set objStream = CreateObject("ADODB.Stream")
                    With objStream
                        .Type = 1
                        .Open
                        .Write objNodeList(i).nextSibling.nodeTypedvalue
                        .SaveToFile pathDir & objNodeList(i).text, 2
                        .Close
                    End With
                    Set objStream = Nothing
                Next
                Set objNodeList = Nothing
            End If
        End If

        Set objXmlFile = Nothing
    End Sub

End Class
%>
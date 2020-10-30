<%
'@title: Class_Ext_UpFile
'@author: 梁无惧
'@date: 2017-02-13
'@description: 无惧上传类v2.2

Class Class_Ext_UpFile

    '@File:上传文件对象集合

    Dim Form, File
    Private allowExt_
    Private noAllowExt_
    Private isDebug_ 
    Private isErr_
    Private errMessage_
    Private oUpFileStream '上传的数据流
    Private isGetData_ '指示是否已执行过GETDATA过程

    '@Version: 版本

    Public Property Get Version
        Version = "无惧上传类 Version V2.0"
    End Property
	
	'@isErr:错误的代码,0或true表示无错

    Public Property Get isErr 
        isErr = isErr_
    End Property
	
    '@errMessage: 错误的字符串信息

    Public Property Get errMessage
        errMessage = errMessage_
    End Property
	
    '@allowExt: 允许上传类型(白名单)

    Public Property Get allowExt 
        allowExt = allowExt_
    End Property

    Public Property Let allowExt(Value)
        allowExt_ = LCase(Value)
    End Property

    '@noAllowExt: 不允许上传类型(黑名单)

    Public Property Get noAllowExt
        noAllowExt = noAllowExt_
    End Property
	
    Public Property Let noAllowExt(Value) 
        noAllowExt_ = LCase(Value)
    End Property
	
	'@isDebug: 是否设置为调试模式

    Public Property Let isDebug(Value)
        isDebug_ = Value
    End Property

    Private Sub Class_Initialize
		Set Form = Server.CreateObject ("Scripting.Dictionary")
		    Form.CompareMode = 1
		Set File = Server.CreateObject ("Scripting.Dictionary")
		    File.CompareMode = 1
        Set oUpFileStream = Server.CreateObject ("ADODB.Stream")
            oUpFileStream.Type = 1
            oUpFileStream.Mode = 3
            oUpFileStream.Open
		'

        noAllowExt = LCase("") '黑名单,可以在这里预设不可上传的文件类型,以文件的后缀名来判断,不分大小写,每个每缀名用;号分开,如果黑名单为空,则判断白名单
        allowExt = LCase("")   '白名单,可以在这里预设可上传的文件类型,以文件的后缀名来判断,不分大小写,每个后缀名用;号分开
        isErr_ = 0
		If IsEmpty(DEBUGS) Then
		   isDebug_ = False
		Else
		   isDebug_ = DEBUGS
		End If
    End Sub

    Private Sub Class_Terminate
		oUpFileStream.Close
		Set oUpFileStream = Nothing
		'
        Form.RemoveAll
        Set Form = Nothing
		'
        File.RemoveAll
        Set File = Nothing
    End Sub

    '@GetData (ByRef MaxSize): 分析上传的数据

    Public Sub GetData (ByRef MaxSize)
        '定义变量
        On Error Resume Next
        If isGetData_ = false Then
            Dim RequestBinData, sSpace, bCrLf, sInfo, iInfoStart, iInfoEnd, tStream, iStart, oFileInfo
            Dim sFormValue, sFileName
            Dim iFindStart, iFindEnd
            Dim iFormStart, iFormEnd, sFormName
            '代码开始
            If InStr(Request.ServerVariables("CONTENT_TYPE"), "multipart/form-data") = 0 Then
                isErr_ = 1
                errMessage_ = "非二进制上传方式!"
                OutErr(errMessage_)
                Exit Sub
            End If
            If Request.TotalBytes < 1 Then '如果没有数据上传
                isErr_ = 1
                errMessage_ = "没有数据上传,这是因为直接提交网址所产生的错误!"
                OutErr(errMessage_)
                Exit Sub
            End If
            If MaxSize > 0 Then '如果限制大小
                If Request.TotalBytes > MaxSize Then
                    isErr_ = 2
                    errMessage_ = "上传的数据超出限制大小!"
                    OutErr(errMessage_)
                    Exit Sub
                End If
            End If

			'将接收到的二进制数据流读取并写入全局oUpFileStream对象
            oUpFileStream.Write Request.BinaryRead (Request.TotalBytes)
            oUpFileStream.Position = 0
            RequestBinData = oUpFileStream.Read '二进制数据保存到RequestBinData变量
            iFormEnd = oUpFileStream.Size
            bCrLf = ChrB (13) & ChrB (10)
            '取得每个项目之间的分隔符
            sSpace = MidB (RequestBinData, 1, InStrB (1, RequestBinData, bCrLf) -1)
            iStart = LenB (sSpace)
            iFormStart = iStart + 2

            '分解项目
			Dim j:j=0
			Set tStream = Server.CreateObject ("ADODB.Stream")
            Do
                iInfoEnd = InStrB (iFormStart, RequestBinData, bCrLf & bCrLf) + 3
                tStream.Type = 1
                tStream.Mode = 3
                tStream.Open
                        oUpFileStream.Position = iFormStart
                        oUpFileStream.CopyTo tStream, iInfoEnd - iFormStart
                tStream.Position = 0
                tStream.Type = 2
                tStream.CharSet = "utf-8"
                sInfo = tStream.ReadText
                '取得表单项目名称
                iFormStart = InStrB (iInfoEnd, RequestBinData, sSpace) -1
                iFindStart = InStr (22, sInfo, "name=""", 1) + 6
                iFindEnd = InStr (iFindStart, sInfo, """", 1)
                sFormName = Mid(sinfo, iFindStart, iFindEnd - iFindStart)
                '如果是文件
                If InStr (45, sInfo, "filename=""", 1) > 0 Then
                    Set oFileInfo = New Class_FileInfo
                    '取得文件属性
                    iFindStart = InStr (iFindEnd, sInfo, "filename=""", 1) + 10
                    iFindEnd = InStr (iFindStart, sInfo, """"&vbCrLf, 1)
                    sFileName = Trim(Mid(sinfo, iFindStart, iFindEnd - iFindStart))
                    oFileInfo.FileName = GetFileName(sFileName)
                    oFileInfo.FilePath = GetFilePath(sFileName)
                    oFileInfo.FileExt = GetFileExt(sFileName)
                    iFindStart = InStr (iFindEnd, sInfo, "Content-Type: ", 1) + 14
                    iFindEnd = InStr (iFindStart, sInfo, vbCr)
                    oFileInfo.FileMIME = Mid(sinfo, iFindStart, iFindEnd - iFindStart)
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
                    tStream.CharSet = "utf-8"
                    sFormValue = tStream.ReadText
                    If Form.Exists (sFormName) Then
                        Form (sFormName) = Form (sFormName) & ", " & sFormValue
                    Else
                        Form.Add sFormName, sFormValue
                    End If
                End If
                tStream.Close
                iFormStart = iFormStart + iStart + 2
                '如果到文件尾了就退出
				if j=0 then
				    j=iFormStart
				elseif j=iFormStart  then
				    isErr_ = 1
					errMessage_ = "火狐上传刷新Bug,避免死循环"
                    OutErr(errMessage_)
					response.end
				end if
            Loop Until (iFormStart + 2) >= iFormEnd
			Set tStream = Nothing
			'
			RequestBinData = ""
            isGetData_ = true
			'
            If Err.Number<>0 Then 
			   errMessage_ = "分解上传数据时发生错误,可能客户端的上传数据不正确或不符合上传数据规则"
			   OutErr(errMessage_)
			end if
        End If
    End Sub

    '@SaveToFile(ByRef Item,ByRef Path): 保存到文件,自动覆盖已存在的同名文件

    Public Function SaveToFile(ByRef Item,ByRef Path)
        SaveToFile = SaveToFileEx(Item, Path, True)
    End Function

    '@AutoSave(Item, Path): 保存到文件,自动设置文件名

    Public Function AutoSave(ByRef Item,ByRef Path)
        AutoSave = SaveToFileEx(Item, Path, false)
    End Function

    '保存到文件,OVER为真时,自动覆盖已存在的同名文件,否则自动把文件改名保存

    Private Function SaveToFileEx(ByRef Item,ByRef Path,ByRef Over)
        On Error Resume Next
        Dim FileExt
        If File.Exists(Item) Then
            Dim oFileStream
            isErr_ = 0
            Set oFileStream = CreateObject ("ADODB.Stream")
            oFileStream.Type = 1
            oFileStream.Mode = 3
            oFileStream.Open
            oUpFileStream.Position = File(Item).FileStart
            oUpFileStream.CopyTo oFileStream, File(Item).FileSize

            FileExt = GetFileExt(Path)
            '
            If isErr_ = 0 Then
                If Over Then
                    If isAllowExt(FileExt) Then
                        oFileStream.SaveToFile Path, 2
                        If Err.Number<>0 Then OutErr("保存文件时出错,请检查路径,是否存在该上传目录!该文件保存路径为" & Path)
                    Else
                        isErr_ = 3
                        errMessage_ = "该后缀名的文件不允许上传"
                        OutErr(errMessage_)
                    End If
                Else
                    Path = GetFilePath(Path)
                    Dim fori, tmpPath
                    fori = 1
                    If isAllowExt(File(Item).FileExt) Then
                        Do
                            fori = fori + 1
                            Err.Clear()
                            tmpPath = Path&GetNewFileName()&"."&File(Item).FileExt
                            oFileStream.SaveToFile tmpPath
                        Loop Until ((Err.Number = 0) Or (fori>50))
                        If Err.Number<>0 Then OutErr("自动保存文件出错,已经测试50次不同的文件名来保存,请检查目录是否存在!该文件最后一次保存时全路径为"&Path&GetNewFileName()&"."&File(Item).FileExt)
                    Else
                        isErr_ = 3
                        errMessage_ = "该后缀名的文件不允许上传"
                        OutErr(errMessage_)
                    End If
                End If
            End If
            oFileStream.Close
            Set oFileStream = Nothing
        Else
            errMessage_ = "不存在该对象(如该文件没有上传,文件为空)!"
            OutErr(errMessage_)
        End If
        If isErr_ = 3 Then SaveToFileEx = "" Else SaveToFileEx = GetFileName(tmpPath)
    End Function

    '取得文件数据

    Public Function FileData(ByRef Item)
        isErr_ = 0
        If File.Exists(Item) Then
            If isAllowExt(File(Item).FileExt) Then
                oUpFileStream.Position = File(Item).FileStart
                FileData = oUpFileStream.Read (File(Item).FileSize)
            Else
                isErr_ = 3
                errMessage_ = "该后缀名的文件不允许上传"
                OutErr(errMessage_)
                FileData = ""
            End If
        Else
            errMessage_ = "不存在该对象(如该文件没有上传,文件为空)!"
            OutErr(errMessage_)
        End If
    End Function


    '取得文件路径

    Public Function GetFilePath(ByRef FullPath)
        If FullPath <> "" Then
            GetFilePath = Left(FullPath, InStrRev(FullPath, "\"))
        Else
            GetFilePath = ""
        End If
    End Function

    '取得文件名

    Public Function GetFileName(ByRef FullPath)
        If FullPath <> "" Then
            GetFileName = Mid(FullPath, InStrRev(FullPath, "\") + 1)
        Else
            GetFileName = ""
        End If
    End Function

    '取得文件的后缀名

    Public Function GetFileExt(ByRef FullPath)
        If FullPath <> "" Then
            GetFileExt = LCase(Mid(FullPath, InStrRev(FullPath, ".") + 1))
        Else
            GetFileExt = ""
        End If
    End Function

    '取得一个不重复的序号

    Public Function GetNewFileName()
        Dim ranNum
        Dim dtNow
        dtNow = Now()
        Randomize
        ranNum = Int(90000 * Rnd) + 10000
        '以下这段由webboy提供
        GetNewFileName = Year(dtNow) & Right("0" & Month(dtNow), 2) & Right("0" & Day(dtNow), 2) & Right("0" & Hour(dtNow), 2) & Right("0" & Minute(dtNow), 2) & Right("0" & Second(dtNow), 2) & ranNum
    End Function

    Public Function isAllowExt(ByRef Ext)
        If noAllowExt = "" Then
            isAllowExt = CBool(InStr(1, ";"&allowExt&";", LCase(";"&Ext&";")))
        Else
            isAllowExt = Not CBool(InStr(1, ";"&noAllowExt&";", LCase(";"&Ext&";")))
        End If
    End Function

    '错误提示

    Public Sub OutErr(ByRef ErrMsg)
        If isDebug_ = true Then
            Response.charset = "utf-8"
            Response.Write ErrMsg
            Response.End
        End If
    End Sub

End Class

'文件属性类

Class Class_FileInfo
    Dim FormName, FileName, FilePath, FileSize, FileMIME, FileStart, FileExt
End Class
%>
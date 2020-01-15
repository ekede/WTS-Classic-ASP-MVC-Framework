<%
'@title: Class_Ext_Wia
'@author: ekede.com
'@date: 2017-02-13
'@description: 扫描仪实现图片缩放类

Class Class_Ext_Wia
    '
    Private v,thumb,img,ip
    Private isDebug_
    Private buildWidth, buildHeight '目标尺寸
    Private csText_, csImg_ '水印

    '@Version: 版本
	
    Public Property Get Version
        version = "1.0"
    End Property

    '@isDebug: 是否设置为调试模式

    Public Property Let isDebug(Value)
        isDebug_ = Value
    End Property
	
    '@csText: 水印文字

    Public Property Let csText(Value)
        csText_ = Value
    End Property
	
    '@csImg: 水印图片
	
    Public Property Let csImg(Value)
        csImg_ = Value
    End Property

    Private Sub Class_Initialize
		If IsEmpty(DEBUGS) Then
		   isDebug_ = False
		Else
		   isDebug_ = DEBUGS
		End If
        '
        On Error Resume Next
        Set v = CreateObject("WIA.Vector") 
		Set thumb = CreateObject("WIA.ImageFile") 
		Set Img = server.CreateObject("WIA.ImageFile")
        Set IP = server.CreateObject("WIA.ImageProcess")
		'
        If Err.Number <> 0 Then OutErr("创建WIA失败")
    End Sub

    Private Sub Class_Terminate
        Set IP = Nothing
        Set Img = Nothing
		Set thumb = Nothing
        Set v = Nothing
    End Sub

    '@BuildPic(ByVal originalPath, Byval buildBasePath, Byval maxWidth, Byval maxHeight): 生成图片

    Public Function BuildPic(ByVal originalPath, Byval buildBasePath, Byval maxWidth, Byval maxHeight)
        On Error Resume Next
		Dim i:i=0
        If originalPath = "" Then Exit Function
        '加载图片
		Img.LoadFile Server.MapPath(originalPath)
		'EXIF过滤器：写一个新的标题标签图像
		If  csText_ <> "" Then
			i=i+1
			IP.Filters.Add IP.FilterInfos("Exif").FilterID
			IP.Filters(i).Properties("ID") = 40091
			IP.Filters(i).Properties("Type") = 1101 'VectorOfBytesImagePropertyType
			v.SetFromString csText_
			IP.Filters(i).Properties("Value") = v
		End If
		'ARGB过滤器：创建一个修改版本的图片
		If  False Then
			i=i+1
			Set c = Img.ARGBData
			For j = 1 To c.Count Step 21 
				c(j) = &HFFFF00FF 'opaque pink (A=255,R=255,G=0,B=255) 
			Next 
			IP.Filters.Add IP.FilterInfos("ARGB").FilterID 
			Set IP.Filters(i).Properties("ARGBData") = c
		End If
        '邮票过滤器：加图片标题信息
		If  csImg_<>"" Then
			i=i+1
			Thumb.LoadFile Server.MapPath(csImg_)
			IP.Filters.Add IP.FilterInfos("Stamp").FilterID
			Set IP.Filters(i).Properties("ImageFile") = Thumb
			IP.Filters(i).Properties("Left") = Img.Width - Thumb.Width
			IP.Filters(i).Properties("Top") = Img.Height - Thumb.Height
		End If
		'裁剪滤镜：裁剪图片
		If  False Then
			i=i+1
			IP.Filters.Add IP.FilterInfos("Crop").FilterID
			IP.Filters(i).Properties("Left") = Img.Width \ 4
			IP.Filters(i).Properties("Top") = Img.Height \ 4
			IP.Filters(i).Properties("Right") = Img.Width \ 4
			IP.Filters(i).Properties("Bottom") = Img.Height \ 4
		End If
		'缩放滤镜：调整图像的大小
		If  True Then
			i=i+1
			IP.Filters.Add IP.FilterInfos("Scale").FilterID
			IP.Filters(i).Properties("MaximumWidth") = maxWidth
			IP.Filters(i).Properties("MaximumHeight") = maxHeight
		End If
		'旋转过滤器：旋转图片
		If  False Then
			i=i+1
			IP.Filters.Add IP.FilterInfos("RotateFlip").FilterID
			IP.Filters(i).Properties("RotationAngle") = 90
		End If
        '图片格式转换：创建一个压缩的JPEG文件
		If  False Then
		    i=i+1
            'wiaFormatBMP  = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
			'wiaFormatPNG  = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
			'wiaFormatGIF  = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
			'wiaFormatTIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
			wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
            IP.Filters.Add IP.FilterInfos("Convert").FilterID 
			IP.Filters(i).Properties("FormatID").Value = wiaFormatJPEG 
			IP.Filters(i).Properties("Quality").Value = 8
		End If
		'最终执行
		Set Img = IP.Apply(Img)
		'保存
        buildFileName = MakeName(originalPath,maxWidth, maxHeight)
		DeleteFile Server.MapPath(buildBasePath)&"\"&buildFileName
        Img.SaveFile Server.MapPath(buildBasePath)&"\"&buildFileName
        '文件名
        If Right(buildBasePath, 1) <> "/" Then buildBasePath = buildBasePath & "/"
		If Err.Number <> 0 then
		   OutErr("缩略图存盘失败,BuildBasePath")
           BuildPic = originalPath
		else
           BuildPic = buildBasePath&buildFileName
		end if
    End Function

    '命名图片

    Private Function MakeName(Byval originalPath, Byval maxWidth, Byval maxHeight)
        Dim pos, oName, oExt
        pos = InStrRev(originalPath, "/") + 1
        oName = Mid(originalPath, pos)
        pos = InStrRev(oName, ".")
        oExt = Mid(oName, pos)
		MakeName = Replace(oName, oExt, "."&maxWidth&"x"&maxHeight&oExt)
    End Function
	
	'删除已存在图片
	
	Private Function DeleteFile(Byval path)
		Dim fso
		Set fso=Server.CreateObject("Scripting.FileSystemObject") 
		If fso.FileExists(path) Then
		   fso.DeleteFile(path)
		End If
		Set fso=Nothing
	End Function

    '错误提示

    Private Sub OutErr(str)
        If isDebug_ Then
            Response.charset = "utf-8"
            Response.Write str
            Response.End
	    End If
    End Sub

End Class
%>
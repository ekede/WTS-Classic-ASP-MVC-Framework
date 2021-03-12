<%
'@title: Class_Ext_Jpeg
'@author: ekede.com
'@date: 2017-02-13
'@description: 图片缩放类

Class Class_Ext_Jpeg
    '
    Private aspJpeg
    Private version_, expires_
    Private isDebug_
    Private buildWidth, buildHeight '目标尺寸
    Private csText_, csImg_ '水印

    '@Version: Persits.Jpeg版本
	
    Public Property Get Version
        version = version_ &" - "& expires_
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
        Set aspJpeg = Server.CreateObject("Persits.Jpeg")
        version_ = aspJpeg.Version
        expires_ = aspJpeg.expires
		'
        If Err.Number <> 0 Then OutErr("AspJpeg组件创建失败")
		If expires_<>"9999-9-9" or expires_<>"9999/9/9" Then OutErr("AspJpeg组件没有注册")
    End Sub

    Private Sub Class_Terminate
        Set aspJpeg = Nothing
    End Sub

    '@BuildPic(ByRef originalPath, ByRef buildBasePath, ByRef maxWidth, ByRef maxHeight): 生成图片

    Public Function BuildPic(ByRef originalPath, ByRef buildBasePath, ByRef maxWidth, ByRef maxHeight)
        On Error Resume Next
        If originalPath = "" Then Exit Function

        aspJpeg.Open Server.MapPath(originalPath)
        If Err.Number <> 0 Then OutErr("原图不存在,OriginalPath")
        '
        ReSize aspJpeg.Width, aspJpeg.Height, maxWidth, maxHeight
        aspJpeg.Width = buildWidth
        aspJpeg.Height = buildHeight
        '
        CanvasText csText_
        CanvasImage csImg_
        '
        buildFileName = MakeName(originalPath,maxWidth, maxHeight)
        aspJpeg.Quality = 100
        aspJpeg.Save Server.MapPath(buildBasePath)&"\"&buildFileName

        '文件名
        If Right(buildBasePath, 1) <> "/" Then buildBasePath = buildBasePath & "/"
		If Err.Number <> 0 then
		   OutErr("缩略图存盘失败,BuildBasePath")
           BuildPic = originalPath
		else
           BuildPic = buildBasePath&buildFileName
		end if
    End Function

    '水印文字

    Private Sub CanvasText(ByRef text)
        If text = "" Then Exit Sub

        Dim x, y
        x = buildWidth -200 '水印横坐标
        y = buildHeight -50 '水印纵坐标
        aspJpeg.Canvas.Font.Size = 12
        aspJpeg.Canvas.Font.Color = &HFFFFFF '颜色
        aspJpeg.Canvas.Font.Bold = True '加粗
        aspJpeg.Canvas.Font.Family = "Aria" '字体
        'aspJpeg.Canvas.Font.Quality = 100            '清晰度
        'aspJpeg.Canvas.Font.ShadowXoffset = 2        '水印文字阴影向右偏移的像素值，输入负值则向左偏
        'aspJpeg.Canvas.Font.ShadowYoffset = 2        '水印文字阴影向下偏移的像素值，输入负值则向右偏
        'aspJpeg.Canvas.Font.ShadowColor = &h0FFFFF   '阴影颜色
        aspJpeg.Canvas.Print x, y, text
    End Sub

    '水印图片

    Private Sub CanvasImage(ByRef pic)
        On Error Resume Next
        If pic = "" Then Exit Sub
        Dim x, y, jpeg2
        Set jpeg2 = server.CreateObject("persits.jpeg")
        jpeg2.Open server.mappath(pic)
        If Err.Number <> 0 Then OutErr("水印图不存在,csImg_")

        x = 1
        y = 1  
        aspJpeg.canvas.drawimage x, y, jpeg2, 0.4, &HFFFFFF 'x,y,水印图,透明度,抽取颜色
        Set jpeg2 = Nothing
    End Sub

    '命名图片

    Private Function MakeName(ByRef originalPath, ByRef maxWidth, ByRef maxHeight)
        Dim pos, oName, oExt
        pos = InStrRev(originalPath, "/") + 1
        oName = Mid(originalPath, pos)
        pos = InStrRev(oName, ".")
        oExt = Mid(oName, pos)
		MakeName = Replace(oName, oExt, "."&maxWidth&"x"&maxHeight&oExt)
    End Function

    '尺寸计算

    Private Sub ReSize(ByRef originalWidth, ByRef originalHeight, ByRef maxWidth, ByRef maxHeight)
        Dim div1, div2
        Dim n1, n2
        div1 = originalWidth / originalHeight
        div2 = originalHeight / originalWidth
        n1 = 0
        n2 = 0
        If originalWidth > maxWidth Then
            n1 = originalWidth / maxWidth
        Else
            buildWidth = originalWidth
        End If
        If originalHeight > maxHeight Then
            n2 = originalHeight / maxHeight
        Else
            buildHeight = originalHeight
        End If
        If n1 <> 0 Or n2 <> 0 Then
            If n1 > n2 Then
                buildWidth = maxWidth
                buildHeight = maxWidth * div2
            Else
                buildWidth = maxHeight * div1
                buildHeight = maxHeight
            End If
        End If
    End Sub

    '错误提示

    Private Sub OutErr(ByRef str)
        If isDebug_ Then
            Response.charset = "utf-8"
            Response.Write str
            Response.End
	    End If
    End Sub

End Class
%>
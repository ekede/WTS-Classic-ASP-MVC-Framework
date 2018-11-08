<%
'@title: Control_Pic
'@author: ekede.com
'@date: 2018-06-17
'@description: 图片及静态文件操作

Class Control_Pic

	'@Index_Action(): 生成并访问缩略图

    Public Function Index_Action()
        '
        pic_path = wts.requests.querystr("p_path")
        pic_name = wts.requests.querystr("p_name")
        pic_width = wts.requests.querystr("p_width")
        pic_height = wts.requests.querystr("p_height")
        pic_ext = wts.requests.querystr("p_ext")
        '
        If pic_name = "" Then Pic_404()
        '
        pic_ori_name = PATH_ROOT&PATH_PIC&PATH_PIC_IMAGES&pic_path&pic_name&"."&pic_ext
        pic_target_path = PATH_ROOT&PATH_PIC&PATH_PIC_THUMBS&pic_path
        '
        If wts.fso.GetRealPath(pic_ori_name)<> -1 Then
            wts.fso.createFolders wts.fso.GetMapPath(pic_target_path)
            pic_url = BuildThumbPic(pic_ori_name, pic_target_path, CInt(pic_width), CInt(pic_height), "")
        Else
            pic_404()
        End If
        '
        If Left(pic_url, 5) = "Error" Or pic_url = "" Then
            pic_404()
        Else
            pic_stream = wts.fso.Reads(wts.fso.GetMapPath(pic_url),"") '读二进制图片
			wts.responses.setContentType = wts.responses.GetContentType("gif")
            wts.responses.SetOutput pic_stream '输出二进制图片
        End If

    End Function
	
	'@Static_Action(): 拷贝并访问静态文件

    Public Function Static_Action()
		Dim APP_MODULE,isFind
		isFind=False
        '
        pic_module = wts.requests.querystr("p_module")
        pic_view = wts.requests.querystr("p_view")
        pic_path = wts.requests.querystr("p_path")
        pic_name = wts.requests.querystr("p_name")
        pic_ext = wts.requests.querystr("p_ext")
		pic_name = pic_name&"."&pic_ext
		'取得APP_MODULE地址
		If PATH_APP <> "" Then
		   If Left(PATH_MODULE,Len(PATH_INC))= PATH_INC Then
		      APP_MODULE = PATH_APP&Right(PATH_MODULE,Len(PATH_MODULE)-Len(PATH_INC))
		   Else
		      APP_MODULE = PATH_APP&PATH_MODULE
		   End If
		End If
		'查看APP_MODULE静态文件是否存在
		If  APP_MODULE <> "" Then
			pic_ori_name = PATH_ROOT&APP_MODULE&pic_module&"/"&PATH_VIEW&pic_view&"/"&pic_path&pic_name
			pic_target_path = PATH_ROOT&PATH_STATIC&pic_module&"/"&pic_view&"/"&pic_path
			'
			If wts.fso.GetRealPath(pic_ori_name)<> -1 Then
				wts.fso.CreateFolders wts.fso.GetMapPath(pic_target_path)
				wts.fso.CopyAFile wts.fso.GetMapPath(pic_ori_name), wts.fso.GetMapPath(pic_target_path&pic_name)
				wts.responses.setContentType = wts.responses.GetContentType(pic_ext)
				wts.responses.SetOutput wts.fso.Reads(wts.fso.GetMapPath(pic_target_path&pic_name),"")
				isFind = true
			Else '当前模板不存在的情况下,查看默认模板是否存在
				pic_def_name = PATH_ROOT&APP_MODULE&pic_module&"/"&PATH_VIEW&wts.site.tplDefaultPath&"/"&pic_path&pic_name
				If wts.fso.GetRealPath(pic_def_name)<> -1 and pic_ori_name <> pic_def_name Then
					wts.fso.CreateFolders wts.fso.GetMapPath(pic_target_path)
					wts.fso.CopyAFile wts.fso.GetMapPath(pic_def_name), wts.fso.GetMapPath(pic_target_path&pic_name)
					wts.responses.setContentType = wts.responses.GetContentType(pic_ext)
					wts.responses.SetOutput wts.fso.Reads(wts.fso.GetMapPath(pic_target_path&pic_name),"")
					isFind = True
				End If
			End If
		End If
		'查看PATH_MODULE静态文件是否存在
		If  isFind = False Then
			pic_ori_name = PATH_ROOT&PATH_MODULE&pic_module&"/"&PATH_VIEW&pic_view&"/"&pic_path&pic_name
			pic_target_path = PATH_ROOT&PATH_STATIC&pic_module&"/"&pic_view&"/"&pic_path
			'
			If wts.fso.GetRealPath(pic_ori_name)<> -1 Then
				wts.fso.CreateFolders wts.fso.GetMapPath(pic_target_path)
				wts.fso.CopyAFile wts.fso.GetMapPath(pic_ori_name), wts.fso.GetMapPath(pic_target_path&pic_name)
				wts.responses.setContentType = wts.responses.GetContentType(pic_ext)
				wts.responses.SetOutput wts.fso.Reads(wts.fso.GetMapPath(pic_target_path&pic_name),"")
			Else '当前模板不存在的情况下,查看默认模板是否存在
				pic_def_name = PATH_ROOT&PATH_MODULE&pic_module&"/"&PATH_VIEW&wts.site.tplDefaultPath&"/"&pic_path&pic_name
				If wts.fso.GetRealPath(pic_def_name)<> -1 and pic_ori_name <> pic_def_name Then
					wts.fso.CreateFolders wts.fso.GetMapPath(pic_target_path)
					wts.fso.CopyAFile wts.fso.GetMapPath(pic_def_name), wts.fso.GetMapPath(pic_target_path&pic_name)
					wts.responses.setContentType = wts.responses.GetContentType(pic_ext)
					wts.responses.SetOutput wts.fso.Reads(wts.fso.GetMapPath(pic_target_path&pic_name),"")
				Else
					wts.errs.AddMsg "no static file"
					wts.errs.Out 404
				End If
			End If
		End If
    End Function

    '404

    Private Sub Pic_404()
        wts.Responses.setStatus = wts.responses.getStatus(404)
        wts.responses.setContentType = "image/gif"
        wts.responses.Transfer PATH_ROOT&PATH_PIC&PATH_PIC_IMAGES&"no.gif"
    End Sub

    '缩图函数

    Private Function BuildThumbPic(originalPath, buildBasePath, maxWidth, maxHeight, Canvas)
	   '#生成缩略图:
        Set jpeg = loader.LoadClass("Ext/jpeg")
	   'jpeg.version
       'jpeg.csText="EKEDE"
       'jpeg.csImg=PATH_PIC&PATH_PIC_IMAGES&"watermark.gif"
        BuildThumbPic = jpeg.BuildPic(originalPath, buildBasePath, maxWidth, maxHeight)
        Set jpeg = Nothing
		'##
    End Function

End Class
%>
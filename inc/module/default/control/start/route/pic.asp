<%
'@title: Control_Start_Route_Pic
'@author: ekede.com
'@date: 2018-06-09
'@description: 图片路由

Class Control_Start_Route_Pic

    Private route_
    Private regEx

    '@route: route对象依赖

    Public Property Let route(Values)
        Set route_ = Values
    End Property
	
    Private Sub Class_Initialize()
        Set regEx = New RegExp
    End Sub

    Private Sub Class_Terminate()
        Set regEx = Nothing
    End Sub
	
   '@DeWrite(ByVal r_path): 解码
   
	Public Sub DeWrite(ByVal r_path)
	    '解图片
	    DeWrite_Static r_path
		If route_.dewrite_on Then Exit Sub
        '解静态
		DeWrite_Pic r_path
	End Sub

    'DeWrite_Pic

    Private Sub DeWrite_Pic(ByVal r_path)
	    if r_path="" Then Exit Sub
        regEx.Pattern = "^"&Replace(PATH_PIC&PATH_PIC_THUMBS, "/", "\/")&"([^\s]+\/|)([^\/]+)\.([0-9]+)x([0-9]+)\.(jpg|png|gif|bmp)(\?.*|)$" ' 设置模式。
		Set matches = regEx.Execute(r_path)
			If matches.Count>0 Then
				route_.requests.querystr("p_path") = matches(0).SubMatches(0)
				route_.requests.querystr("p_name") = matches(0).SubMatches(1)
				route_.requests.querystr("p_width") = matches(0).SubMatches(2)
				route_.requests.querystr("p_height") = matches(0).SubMatches(3)
				route_.requests.querystr("p_ext") = matches(0).SubMatches(4)
				'
				c = "pic"
				If  route_.loader.LoadFile(PATH_MODULE&route_.module&"/"&PATH_CONTROL&c&".asp")<> -1 Then
					route_.control = c
					route_.action = "index"
					route_.dewrite_on = true
				End If
			Else
				pageext = route_.fun.GetExt(r_path)
				If pageext = ".gif" Or pageext = ".jpg" Or pageext = ".png" Or pageext = ".bmp" Then
					c = "pic"
					If route_.loader.LoadFile(PATH_MODULE&module&"/"&PATH_CONTROL&c&".asp")<> -1 Then
						route_.control = c
						route_.action = "index"
						route_.dewrite_on = true
					End If
				End If
			End If
        Set matches = nothing
    End Sub
	
	'DeWrite_Static
	
    Private Sub DeWrite_Static(ByVal r_path)
	    if r_path="" Then Exit Sub
        regEx.Pattern = "^"&PATH_STATIC&"([^\/]+)\/([^\/]+)\/([^\s]+\/|)([^\/]+)\.(css|js|jpg|gif|png|bmp|svg|ico|woff2|otf|ttf|eot)(\?.*|)$" ' 设置模式。
		Set matches = regEx.Execute(r_path)
			If  matches.Count>0 Then
				route_.requests.querystr("p_module") = matches(0).SubMatches(0)
				route_.requests.querystr("p_view") = matches(0).SubMatches(1)
				route_.requests.querystr("p_path") = matches(0).SubMatches(2)
				route_.requests.querystr("p_name") = matches(0).SubMatches(3)
			    route_.requests.querystr("p_ext") = matches(0).SubMatches(4)
				c = "pic"
				If  route_.loader.LoadFile(PATH_MODULE&route_.module&"/"&PATH_CONTROL&c&".asp")<> -1 Then
					route_.control = c
					route_.action = "static"
					route_.dewrite_on = true
				End If
			End If
        Set matches = nothing
    End Sub
	
    '@ReWrite(ByVal str): 编码未使用

    Public Function ReWrite(ByVal str)
    End Function
	
    '@ReWritePic(Byref base,Byref pic_name,Byref pic_width,Byref pic_height,Byref picdefault): 编码图片

    Public Function ReWritePic(Byref base,Byref pic_name,Byref pic_width,Byref pic_height,Byref picdefault)
        Dim str, pic_ext
        If IsNull(pic_name) Then pic_name = ""
        If pic_name = "" Then
            If picdefault<>"" Then
                str = picdefault
            Else
                str = "images/no.gif"
            End If
        Else
            str = pic_name
        End If
        str = LCase(str)
        ReWritePic = base&PATH_PIC&ReWrite_P(str, pic_width, pic_height)
    End Function
	
    Private Function ReWrite_P(Byref pic_name,Byref pic_width,Byref pic_height)
        Dim str, pic_ext
        str = pic_name
        If IsNumeric(pic_width) And IsNumeric(pic_height) Then
            str = Replace(str, PATH_PIC_IMAGES, PATH_PIC_THUMBS)
            pic_ext = route_.fun.GetExt(str)
            str = Replace(str, pic_ext, "."&pic_width&"x"&pic_height&pic_ext)
        End If
        ReWrite_P = str
    End Function
	
    '@ReWriteStatic(Byref base,Byref names): 编码static文件
	
    Public Function ReWriteStatic(Byref base,Byref names)
        ReWriteStatic = base&names
    End Function
	
End Class
%>
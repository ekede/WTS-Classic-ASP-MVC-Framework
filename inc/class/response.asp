<%
'@title: Class_Response
'@author: ekede.com
'@date: 2017-12-12
'@description: Response对象

Class Class_Response

    Private output_
    Private headers_, buffer_, charset_, contentType_, status_

    '@setBuffer: 缓存输出

    Public Property Let setBuffer(str)
        buffer_ = str
    End Property
	
    '@setCharset: 编码类型

    Public Property Let setCharset(str)
        charset_ = str
    End Property
	
    '@setContentType: 文档类型

    Public Property Let setContentType(str)
	    ContentType_ = str
    End Property
	
    '@setStatus: http状态

    Public Property Let setStatus(str)
	    status_ = str
    End Property

    Private Sub Class_Initialize()
        Set headers_ = Server.CreateObject("Scripting.Dictionary")
        '
        buffer_ = true
        charset_ = "UTF-8" '列表
        contentType_ = "text/html" '列表
        status_ = "200 ok" '列表
    End Sub

    Private Sub Class_Terminate()
        Set headers_ = Nothing
    End Sub
	
    '@SetHeader(outs): 设置header

    Public Function SetHeader(names, content)
        headers_(names) = content
    End Function
	
    '@SetOutput(outs): 设置内容

    Public Function SetOutput(outs)
        output_ = outs
    End Function

    '@GetOutput(): 查看内容

    Public Function GetOutput()
        GetOutput = output_
    End Function

    '@OutPuts(): 浏览器输出

    Public Sub OutPuts()
        Response.Clear()
        Response.Buffer = buffer_
        Response.ContentType = contentType_
        Response.Status = status_
        For Each keys in headers_
            Response.AddHeader keys, headers_(keys)
        Next
        If TypeName(output_)="String" Then
            Response.CodePage = 65001
            Response.Charset = charset_
            Response.Write output_
        ElseIf TypeName(output_)="Byte()" Then
            Response.BinaryWrite output_
            Response.Flush
        End If
    End Sub

    '@Transfer(path): 转向包含

    Public Sub Transfer(path)
        Response.ContentType = ContentType_
        Response.Status = Status_
        Server.transfer(path)
    End Sub

    '@Direct(Url): 跳转

    Public Sub Direct(Url)
        response.redirect(Replace(Url, "&amp;", "&"))
    End Sub

    '@Direct301(Url): 301跳转

    Public Sub Direct301(Url)
        Response.Status = GetStatus(301)
        Response.AddHeader "Location", Url
        Response.End
    End Sub
	
    '@Die(str): 中断
	
	Public Sub Die(str)
	    Response.Charset = charset_
	    Response.write str
		Response.End
	End Sub
	
    '@GetStatus(n): 根据状态码取http状态字符串
	
    Public Function GetStatus(n)
        Select Case n
            Case 301
                GetStatus = "301 Moved Permanently"
            Case 401
                GetStatus = "404 Unauthorized"
            Case 404
                GetStatus = "404 Not Found"
            Case 500
                GetStatus = "500 Internal Server Error"
            Case Else
                GetStatus = "200 ok"
        End Select
	End Function

    '@GetContentType(ext): 根据扩展名取Content-Type(Mime-Type)字符串
	
    Public Function GetContentType(ext)
		Dim e,s
		e=LCase(replace(ext,".",""))
		Select Case e
			Case "html" 
				s="text/html"
			Case "xhtml" 
				s="text/html"
			Case "xml"  
				s="text/xml"
			Case "xsl"  
				s="text/xml"	 
			Case "xslt" 
				s="text/xml"
			Case "wml"  
				s="text/vnd.wap.wml"
			Case "wsdl" 
				s="text/xml"
			'
			Case "css"  
				s="text/css"
			Case "js"   
				s="application/x-javascript"
			Case "json" 
				s="application/json"
			'
			Case "woff" 
				s="application/x-font-woff"
			Case "woff2" 
				s="application/x-font-woff2"
			Case "otf"  
				s="application/x-font-opentype"
			Case "ttf"  
				s="application/x-font-truetype"
			Case "eot"  
				s="application/vnd.ms-fontobject"
			'
			Case "png" 
				s="image/png"
			Case "jpg"  
				s="image/jpeg"
			Case "gif"  
				s="image/gif"
			Case "bmp"  
				s="application/x-bmp"
			Case "svg"  
				s="image/svg+xml"
			Case "ico"  
				s="application/x-ico"
	        '
			Case "pdf"  
				s="application/pdf"
			Case "xls"  
				s="application/x-xls"
			Case "doc" 	
				s="application/msword"
			Case "ppt" 
				s="application/x-ppt"
			Case "zip" 
				s="application/zip"
			Case "gzip" 
				s="application/gzip"
			'
		    Case Else
				s=""
		End Select
		'
		GetContentType = s
	End Function

End Class
%>